#include "cli_test_support.h"

#include <dirent.h>
#include <fcntl.h>
#include <signal.h>
#include <sys/stat.h>
#include <sys/wait.h>
#include <unistd.h>

#include <algorithm>
#include <cstdint>
#include <cstdio>
#include <cstdlib>
#include <fstream>
#include <iterator>
#include <utility>

#include "c_api/formulon_c.h"
#include "defined_name.h"
#include "support/ooxml_package_fixture.h"

namespace formulon::cli_test {

// Quotes `s` in single quotes for `/bin/sh`, doubling any embedded
// single quotes via `'\''`. Sufficient for the file paths and short
// flag values the tests pass; not a full shell quoter.
std::string sh_quote(const std::string& s) {
  std::string out;
  out.reserve(s.size() + 2);
  out.push_back('\'');
  for (char c : s) {
    if (c == '\'') {
      out.append("'\\''");
    } else {
      out.push_back(c);
    }
  }
  out.push_back('\'');
  return out;
}

PathGuard::PathGuard(std::string p) : path(std::move(p)) {}

PathGuard::~PathGuard() {
  if (!path.empty()) {
    std::remove(path.c_str());
  }
}

DirGuard::DirGuard(std::string p) : path(std::move(p)) {}

DirGuard::~DirGuard() {
  ::chmod(path.c_str(), 0700);
  for (const std::string& entry : list_directory(path)) {
    const std::string full = path + "/" + entry;
    if (std::remove(full.c_str()) != 0) {
      ::rmdir(full.c_str());
    }
  }
  ::rmdir(path.c_str());
}

// Prefix every temp fixture in this file is built from.
//
// The paths are derived from the pid rather than fixed, so two runs of
// this binary -- a `ctest -j` fan-out, or a rerun started while an
// earlier one is still finishing -- never share a fixture and never fail
// on each other's leftovers. `TMPDIR` is honoured where the platform
// sets one, falling back to `/tmp`.
const std::string& temp_prefix() {
  static const std::string prefix = [] {
    const char* configured = std::getenv("TMPDIR");
    std::string dir = (configured != nullptr && configured[0] != '\0') ? configured : "/tmp";
    if (dir.back() != '/') {
      dir.push_back('/');
    }
    return dir + "fm_cli_" + std::to_string(::getpid()) + "_";
  }();
  return prefix;
}

// Absolute path of a temp fixture named `name`, unique to this process.
std::string temp_path(std::string_view name) {
  return temp_prefix() + std::string(name);
}

// Entries of `dir` excluding `.` and `..`, sorted so a comparison against
// an expected listing is stable. An unreadable directory yields an empty
// listing, which reads the same as "nothing was left behind" -- callers
// that care assert on the readable case.
std::vector<std::string> list_directory(const std::string& dir) {
  std::vector<std::string> entries;
  DIR* handle = ::opendir(dir.c_str());
  if (handle == nullptr) {
    return entries;
  }
  while (const dirent* entry = ::readdir(handle)) {
    const std::string name(entry->d_name);
    if (name != "." && name != "..") {
      entries.push_back(name);
    }
  }
  ::closedir(handle);
  std::sort(entries.begin(), entries.end());
  return entries;
}

// Spawns the CLI with `args` (each argument is sh-quoted before being
// joined into a single command line). When `merge_streams` is true,
// stderr is folded into stdout via `2>&1`; otherwise stderr is
// captured to a temp file via `2>` and read back. The temp-file path
// avoids races with shell escaping.
CliRun run_cli(const std::vector<std::string>& args, bool merge_streams, bool close_stdout,
               const std::string& stdin_path) {
  CliRun out;
  std::string cmd = sh_quote(FORMULON_CLI_PATH);
  for (const auto& a : args) {
    cmd.push_back(' ');
    cmd.append(sh_quote(a));
  }
  if (!stdin_path.empty()) {
    cmd.append(" <");
    cmd.append(sh_quote(stdin_path));
  }
  // Capture stderr separately to a temp file so we can assert on it
  // without ambiguity. The shell's `>` redirect is portable on macOS
  // and Linux. We append rather than truncate so the temp file is
  // self-contained per call.
  std::string tmpl = temp_path("stderr_XXXXXX");
  const int fd = mkstemp(tmpl.data());
  if (fd < 0) {
    out.exit_code = -2;
    return out;
  }
  ::close(fd);
  if (close_stdout) {
    cmd.append(" 1>&- 2>");
    cmd.append(sh_quote(tmpl));
  } else if (merge_streams) {
    cmd.append(" 2>&1");
  } else {
    cmd.append(" 2>");
    cmd.append(sh_quote(tmpl));
  }

  FILE* pipe = popen(cmd.c_str(), "r");
  if (pipe == nullptr) {
    out.exit_code = -3;
    std::remove(tmpl.c_str());
    return out;
  }
  char buf[4096];
  std::size_t got = 0;
  while ((got = std::fread(buf, 1, sizeof(buf), pipe)) > 0) {
    out.stdout_text.append(buf, got);
  }
  const int wait_rc = pclose(pipe);
  // `pclose` returns the wait status; on POSIX, `WEXITSTATUS` is in
  // the upper byte. We approximate it with `(wait_rc >> 8) & 0xff`,
  // which matches `/bin/sh` semantics.
  if (wait_rc < 0) {
    out.exit_code = -4;
  } else {
    out.exit_code = (wait_rc >> 8) & 0xff;
  }

  if (!merge_streams) {
    std::ifstream in(tmpl);
    if (in) {
      std::string line;
      while (std::getline(in, line)) {
        out.stderr_text.append(line);
        out.stderr_text.push_back('\n');
      }
    }
  }
  std::remove(tmpl.c_str());
  return out;
}

namespace {

std::string slurp(const std::string& path) {
  std::ifstream in(path, std::ios::binary);
  return std::string((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
}

// Opens the slave side of a fresh pseudo-terminal and keeps the master open
// for the lifetime of the object, so the slave stays a live terminal.
struct Pty {
  int master = -1;
  int slave = -1;
  Pty() {
    master = ::posix_openpt(O_RDWR | O_NOCTTY);
    if (master < 0 || ::grantpt(master) != 0 || ::unlockpt(master) != 0) {
      return;
    }
    const char* name = ::ptsname(master);
    if (name != nullptr) {
      slave = ::open(name, O_RDWR | O_NOCTTY);
    }
  }
  Pty(const Pty&) = delete;
  Pty& operator=(const Pty&) = delete;
  ~Pty() {
    if (slave >= 0) {
      ::close(slave);
    }
    if (master >= 0) {
      ::close(master);
    }
  }
  bool ok() const { return slave >= 0; }
};

}  // namespace

CliRun run_cli_spawn(const std::vector<std::string>& args, const SpawnOptions& options) {
  CliRun out;
  const std::string out_path = temp_path("spawn_out_" + std::to_string(::getpid()) + "_" +
                                         std::to_string(reinterpret_cast<std::uintptr_t>(&out)));
  const std::string err_path = out_path + ".err";
  PathGuard out_guard(out_path);
  PathGuard err_guard(err_path);

  Pty in_pty;
  Pty out_pty;
  if ((options.stdin_tty && !in_pty.ok()) || (options.stdout_tty && !out_pty.ok())) {
    out.exit_code = -2;
    return out;
  }
  int broken[2] = {-1, -1};
  if (options.stdout_broken_pipe) {
    if (::pipe(broken) != 0) {
      out.exit_code = -2;
      return out;
    }
    ::close(broken[0]);
  }

  std::vector<std::string> storage;
  storage.emplace_back(FORMULON_CLI_PATH);
  storage.insert(storage.end(), args.begin(), args.end());
  std::vector<char*> argv;
  for (std::string& s : storage) {
    argv.push_back(s.data());
  }
  argv.push_back(nullptr);

  const pid_t pid = ::fork();
  if (pid < 0) {
    out.exit_code = -3;
    return out;
  }
  if (pid == 0) {
    const int devnull = ::open("/dev/null", O_RDONLY);
    const int out_file = ::open(out_path.c_str(), O_WRONLY | O_CREAT | O_TRUNC, 0600);
    const int err_file = ::open(err_path.c_str(), O_WRONLY | O_CREAT | O_TRUNC, 0600);
    ::dup2(options.stdin_tty ? in_pty.slave : devnull, 0);
    if (options.stdout_tty) {
      ::dup2(out_pty.slave, 1);
    } else if (options.stdout_broken_pipe) {
      ::dup2(broken[1], 1);
    } else {
      ::dup2(out_file, 1);
    }
    ::dup2(err_file, 2);
    ::execv(argv[0], argv.data());
    ::_exit(127);
  }
  if (broken[1] >= 0) {
    ::close(broken[1]);
  }

  constexpr int kPollMillis = 10;
  constexpr int kMaxPolls = 500;
  int status = 0;
  bool reaped = false;
  for (int i = 0; i < kMaxPolls && !reaped; ++i) {
    reaped = ::waitpid(pid, &status, WNOHANG) == pid;
    if (!reaped) {
      ::usleep(kPollMillis * 1000);
    }
  }
  if (!reaped) {
    ::kill(pid, SIGKILL);
    ::waitpid(pid, &status, 0);
    out.exit_code = -5;
  } else {
    out.exit_code = WIFEXITED(status) ? WEXITSTATUS(status) : -6;
  }
  out.stdout_text = slurp(out_path);
  out.stderr_text = slurp(err_path);
  return out;
}

bool write_bytes(const std::string& path, std::string_view bytes) {
  std::ofstream f(path, std::ios::binary | std::ios::trunc);
  f.write(bytes.data(), static_cast<std::streamsize>(bytes.size()));
  return static_cast<bool>(f);
}

// The load-diagnostic warning lines in `text`, with the
// `formulon: <subcommand>: ` prefix stripped so reports emitted by
// different subcommands compare equal. Save-side warnings are left out:
// only the command that writes a workbook produces those.
std::vector<std::string> read_diagnostic_lines(const std::string& text) {
  std::vector<std::string> lines;
  std::size_t start = 0;
  while (start < text.size()) {
    const std::size_t end = text.find('\n', start);
    const std::string line = text.substr(start, end == std::string::npos ? std::string::npos : end - start);
    const std::size_t marker = line.find("warning: ");
    if (marker != std::string::npos && line.find("read diagnostics") != std::string::npos) {
      lines.push_back(line.substr(marker));
    }
    if (end == std::string::npos) {
      break;
    }
    start = end + 1;
  }
  return lines;
}

// Number of descriptors this process currently holds open. Only the
// delta across a call matters, so scanning a fixed low range is enough:
// the tests below open a handful of files and never approach the cap.
int open_descriptor_count() {
  int count = 0;
  for (int fd = 0; fd < 256; ++fd) {
    if (::fcntl(fd, F_GETFD) != -1) {
      ++count;
    }
  }
  return count;
}

// Builds a tiny `.xlsx` workbook in memory via the C ABI and writes
// it to `path`. Used by `recalc` and `dump` tests that need a fixture
// on disk. Returns `true` on success.
bool write_fixture_workbook(const std::string& path, std::string_view extra_text, bool iterative,
                            int32_t max_iterations, double max_change) {
  fm_workbook_t* wb = nullptr;
  if (fm_workbook_create_default(&wb) != 0) {
    return false;
  }
  // A1 = 7 (literal), B1 = =A1+1 (formula).
  if (fm_sheet_set_number(wb, 0, 0, 0, 7.0) != 0) {
    fm_workbook_destroy(wb);
    return false;
  }
  if (fm_sheet_set_formula(wb, 0, 0, 1, "=A1+1") != 0) {
    fm_workbook_destroy(wb);
    return false;
  }
  fm_iterative_settings_patch_t settings = FM_STRUCT_INIT(fm_iterative_settings_patch_t);
  settings.enabled_op = FM_PATCH_OP_SET;
  settings.enabled = 1;
  settings.max_iterations_op = FM_PATCH_OP_SET;
  settings.max_iterations = static_cast<uint32_t>(max_iterations);
  settings.max_change_op = FM_PATCH_OP_SET;
  settings.max_change = max_change;
  if (iterative && fm_workbook_update_iterative(wb, &settings) != 0) {
    fm_workbook_destroy(wb);
    return false;
  }
  if (fm_workbook_recalc(wb) != 0) {
    fm_workbook_destroy(wb);
    return false;
  }
  if (!extra_text.empty() && fm_sheet_set_text(wb, 0, 1, 0, std::string(extra_text).c_str()) != 0) {
    fm_workbook_destroy(wb);
    return false;
  }
  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  if (fm_workbook_save(wb, &bytes, &len) != 0) {
    fm_workbook_destroy(wb);
    return false;
  }
  std::ofstream out(path, std::ios::binary | std::ios::trunc);
  if (!out) {
    fm_buffer_destroy(bytes);
    fm_workbook_destroy(wb);
    return false;
  }
  out.write(reinterpret_cast<const char*>(bytes), static_cast<std::streamsize>(len));
  out.close();
  fm_buffer_destroy(bytes);
  fm_workbook_destroy(wb);
  return out.good();
}

// A wide, pure formula layer makes the CLI's parallel route observable
// without relying on volatile functions or timing. The caller receives an
// ordinary .xlsx that the command must load, recalculate, save, and whose
// last lane we inspect after reopening.
bool write_wide_recalc_fixture(const std::string& path) {
  fm_workbook_t* wb = nullptr;
  if (fm_workbook_create_default(&wb) != 0) {
    return false;
  }
  constexpr std::uint32_t kRows = 48U;
  for (std::uint32_t row = 0U; row < kRows; ++row) {
    const std::string formula = "=A" + std::to_string(row + 1U) + "*2+1";
    if (fm_sheet_set_number(wb, 0U, row, 0U, static_cast<double>(row + 1U)) != 0 ||
        fm_sheet_set_formula(wb, 0U, row, 1U, formula.c_str()) != 0) {
      fm_workbook_destroy(wb);
      return false;
    }
  }
  std::uint8_t* bytes = nullptr;
  std::size_t len = 0U;
  if (fm_workbook_save(wb, &bytes, &len) != 0) {
    fm_workbook_destroy(wb);
    return false;
  }
  std::ofstream out(path, std::ios::binary | std::ios::trunc);
  if (!out) {
    fm_buffer_destroy(bytes);
    fm_workbook_destroy(wb);
    return false;
  }
  out.write(reinterpret_cast<const char*>(bytes), static_cast<std::streamsize>(len));
  out.close();
  fm_buffer_destroy(bytes);
  fm_workbook_destroy(wb);
  return out.good();
}

bool write_out_of_range_defined_names_fixture(const std::string& path) {
  formulon::Workbook wb = formulon::Workbook::create();
  for (std::size_t i = 1; i < 7; ++i) {
    wb.add_sheet("Sheet" + std::to_string(i + 1));
  }

  std::vector<formulon::DefinedName> names;
  names.push_back(formulon::DefinedName{"Before", "1", -1, false, ""});
  names.push_back(formulon::DefinedName{"Bad", "Sheet1!$A$1", 99, false, ""});
  names.push_back(formulon::DefinedName{"After", "2", -1, false, ""});
  wb.set_defined_names(std::move(names));

  auto saved = wb.save();
  if (!saved) {
    return false;
  }
  std::ofstream out(path, std::ios::binary | std::ios::trunc);
  if (!out) {
    return false;
  }
  const std::vector<std::uint8_t>& bytes = saved.value();
  out.write(reinterpret_cast<const char*>(bytes.data()), static_cast<std::streamsize>(bytes.size()));
  out.close();
  return out.good();
}

bool write_lossy_xlsx_fixture(const std::string& path) {
  fm_workbook_t* wb = nullptr;
  if (fm_workbook_create_default(&wb) != 0) {
    return false;
  }
  if (fm_sheet_set_formula(wb, 0, 0, 0, "=T[C]") != 0) {
    fm_workbook_destroy(wb);
    return false;
  }
  fm_hyperlink_t hyperlink = FM_STRUCT_INIT(fm_hyperlink_t);
  hyperlink.target = "https://example.com";
  if (fm_sheet_set_hyperlink(wb, 0, 0, 0, &hyperlink) != 0) {
    fm_workbook_destroy(wb);
    return false;
  }
  fm_cell_rect_t range{0, 0, 0, 0};
  fm_data_validation_t validation = FM_STRUCT_INIT(fm_data_validation_t);
  validation.ranges = &range;
  validation.range_count = 1;
  validation.type = 3;
  validation.formula1 = "$A$1:$A$2";
  std::size_t validation_index = 0;
  if (fm_sheet_add_validation(wb, 0, &validation, &validation_index) != 0 ||
      fm_sheet_set_auto_filter_xml(wb, 0, "<autoFilter ref=\"A1:B2\"/>") != 0) {
    fm_workbook_destroy(wb);
    return false;
  }
  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  if (fm_workbook_save(wb, &bytes, &len) != 0) {
    fm_workbook_destroy(wb);
    return false;
  }
  std::ofstream out(path, std::ios::binary | std::ios::trunc);
  if (!out) {
    fm_buffer_destroy(bytes);
    fm_workbook_destroy(wb);
    return false;
  }
  out.write(reinterpret_cast<const char*>(bytes), static_cast<std::streamsize>(len));
  out.close();
  fm_buffer_destroy(bytes);
  fm_workbook_destroy(wb);
  return out.good();
}

std::vector<std::uint8_t> append_empty_zip_entry(const std::vector<std::uint8_t>& bytes, std::string_view name) {
  const std::uint8_t sig[] = {0x50, 0x4b, 0x05, 0x06};
  const auto it = std::find_end(bytes.begin(), bytes.end(), std::begin(sig), std::end(sig));
  if (it == bytes.end())
    return {};
  const std::size_t eocd = static_cast<std::size_t>(it - bytes.begin());
  const auto read16 = [&](std::size_t p) { return static_cast<std::uint16_t>(bytes[p] | (bytes[p + 1] << 8U)); };
  const auto read32 = [&](std::size_t p) {
    return static_cast<std::uint32_t>(bytes[p] | (bytes[p + 1] << 8U) | (bytes[p + 2] << 16U) | (bytes[p + 3] << 24U));
  };
  const std::uint16_t count = read16(eocd + 10U);
  const std::uint32_t central_size = read32(eocd + 12U);
  const std::uint32_t central_offset = read32(eocd + 16U);
  const auto put16 = [](std::vector<std::uint8_t>& out, std::uint16_t x) {
    out.push_back(static_cast<std::uint8_t>(x));
    out.push_back(static_cast<std::uint8_t>(x >> 8U));
  };
  const auto put32 = [](std::vector<std::uint8_t>& out, std::uint32_t x) {
    out.push_back(static_cast<std::uint8_t>(x));
    out.push_back(static_cast<std::uint8_t>(x >> 8U));
    out.push_back(static_cast<std::uint8_t>(x >> 16U));
    out.push_back(static_cast<std::uint8_t>(x >> 24U));
  };
  const std::string n(name);
  std::vector<std::uint8_t> local;
  put32(local, 0x04034b50U);
  put16(local, 20);
  put16(local, 0);
  put16(local, 0);
  put16(local, 0);
  put16(local, 0);
  put32(local, 0);
  put32(local, 0);
  put32(local, 0);
  put16(local, static_cast<std::uint16_t>(n.size()));
  put16(local, 0);
  local.insert(local.end(), n.begin(), n.end());
  std::vector<std::uint8_t> central;
  put32(central, 0x02014b50U);
  put16(central, 20);
  put16(central, 20);
  put16(central, 0);
  put16(central, 0);
  put16(central, 0);
  put16(central, 0);
  put32(central, 0);
  put32(central, 0);
  put32(central, 0);
  put16(central, static_cast<std::uint16_t>(n.size()));
  put16(central, 0);
  put16(central, 0);
  put16(central, 0);
  put16(central, 0);
  put32(central, 0);
  put32(central, central_offset);
  central.insert(central.end(), n.begin(), n.end());
  std::vector<std::uint8_t> out(bytes.begin(), bytes.begin() + central_offset);
  out.insert(out.end(), local.begin(), local.end());
  out.insert(out.end(), bytes.begin() + central_offset, bytes.begin() + central_offset + central_size);
  out.insert(out.end(), central.begin(), central.end());
  put32(out, 0x06054b50U);
  put16(out, 0);
  put16(out, 0);
  put16(out, count + 1U);
  put16(out, count + 1U);
  put32(out, central_size + static_cast<std::uint32_t>(central.size()));
  put32(out, central_offset + static_cast<std::uint32_t>(local.size()));
  put16(out, 0);
  return out;
}

// An `.xlsx` whose sheet carries one unusable reference per overlay kind and
// whose workbook part declares a content type the reader does not recognise.
// The engine's writer never produces either, so the package is assembled by
// hand.
bool write_lossy_ooxml_fixture(const std::string& path) {
  const std::string content_types = formulon::test::OoxmlContentTypes("application/vnd.formulon-test.unknown+xml");
  const std::vector<std::uint8_t> bytes = formulon::test::BuildOoxmlPackage(
      content_types,
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData/>"
      "<mergeCells><mergeCell ref=\"nope\"/></mergeCells>"
      "<conditionalFormatting><cfRule type=\"expression\" priority=\"1\"/></conditionalFormatting>"
      "</worksheet>");
  if (bytes.empty()) {
    return false;
  }
  std::ofstream out(path, std::ios::binary | std::ios::trunc);
  out.write(reinterpret_cast<const char*>(bytes.data()), static_cast<std::streamsize>(bytes.size()));
  return out.good();
}

bool write_dropped_xlsb_fixture(const std::string& path) {
  fm_workbook_t* wb = nullptr;
  if (fm_workbook_create_default(&wb) != 0)
    return false;
  std::uint8_t* raw = nullptr;
  std::size_t len = 0;
  const bool saved = fm_workbook_save_as(wb, FM_WORKBOOK_FORMAT_XLSB, &raw, &len) == 0;
  fm_workbook_destroy(wb);
  if (!saved)
    return false;
  std::vector<std::uint8_t> fixture =
      append_empty_zip_entry(std::vector<std::uint8_t>(raw, raw + len), "xl/dropped.unknown");
  fm_buffer_destroy(raw);
  if (fixture.empty())
    return false;
  std::ofstream out(path, std::ios::binary | std::ios::trunc);
  out.write(reinterpret_cast<const char*>(fixture.data()), static_cast<std::streamsize>(fixture.size()));
  return out.good();
}

}  // namespace formulon::cli_test
