// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for StructuredLog. The emitter writes to stderr, so each test
// redirects the stderr file descriptor (fd 2) through a pipe, runs the log
// call, then reads the captured bytes back. Only POSIX platforms provide the
// `dup`/`dup2`/`pipe` primitives; the tests are skipped elsewhere.

#include "utils/structured_log.h"

#include <cstdio>
#include <cstring>
#include <string>
#include <string_view>

#include "gtest/gtest.h"
#include "utils/error.h"

#if defined(__unix__) || defined(__APPLE__)
#define FORMULON_HAVE_POSIX_PIPE 1
#include <fcntl.h>
#include <unistd.h>
#else
#define FORMULON_HAVE_POSIX_PIPE 0
#endif

namespace formulon {
namespace {

#if FORMULON_HAVE_POSIX_PIPE

class StderrCapture {
 public:
  StderrCapture() {
    std::fflush(stderr);
    saved_fd_ = ::dup(STDERR_FILENO);
    int pipefd[2] = {-1, -1};
    if (::pipe(pipefd) == 0) {
      read_fd_ = pipefd[0];
      write_fd_ = pipefd[1];
      // Make the read end non-blocking so we never hang in `Read()` even if
      // the emitter wrote less than we expected.
      int flags = ::fcntl(read_fd_, F_GETFL, 0);
      ::fcntl(read_fd_, F_SETFL, flags | O_NONBLOCK);
      ::dup2(write_fd_, STDERR_FILENO);
    }
  }

  ~StderrCapture() { Restore(); }

  StderrCapture(const StderrCapture&) = delete;
  StderrCapture& operator=(const StderrCapture&) = delete;

  std::string Read() {
    std::fflush(stderr);
    Restore();
    std::string out;
    if (read_fd_ < 0)
      return out;
    char buf[1024];
    for (;;) {
      ssize_t n = ::read(read_fd_, buf, sizeof(buf));
      if (n > 0) {
        out.append(buf, static_cast<size_t>(n));
        continue;
      }
      break;
    }
    ::close(read_fd_);
    read_fd_ = -1;
    return out;
  }

 private:
  void Restore() {
    if (saved_fd_ >= 0) {
      std::fflush(stderr);
      ::dup2(saved_fd_, STDERR_FILENO);
      ::close(saved_fd_);
      saved_fd_ = -1;
    }
    if (write_fd_ >= 0) {
      ::close(write_fd_);
      write_fd_ = -1;
    }
  }

  int saved_fd_ = -1;
  int read_fd_ = -1;
  int write_fd_ = -1;
};

bool Contains(std::string_view haystack, std::string_view needle) {
  return haystack.find(needle) != std::string_view::npos;
}

TEST(StructuredLogTest, InfoEmitsEventAndLevel) {
  StderrCapture cap;
  StructuredLog("cell.evaluated").info();
  std::string out = cap.Read();
  EXPECT_TRUE(Contains(out, "\"event\":\"cell.evaluated\"")) << out;
  EXPECT_TRUE(Contains(out, "\"level\":\"info\"")) << out;
}

TEST(StructuredLogTest, FieldsSerialized) {
  StderrCapture cap;
  StructuredLog("op.done")
      .field("sheet", std::string_view("Sheet1"))
      .field("duration_us", static_cast<int64_t>(123))
      .field("ratio", 0.5)
      .field("ok", true)
      .info();
  std::string out = cap.Read();
  EXPECT_TRUE(Contains(out, "\"sheet\":\"Sheet1\"")) << out;
  EXPECT_TRUE(Contains(out, "\"duration_us\":123")) << out;
  EXPECT_TRUE(Contains(out, "\"ratio\":")) << out;
  EXPECT_TRUE(Contains(out, "\"ok\":true")) << out;
}

TEST(StructuredLogTest, ErrorCodeFieldIncluded) {
  StderrCapture cap;
  StructuredLog("parser.error").error_code(FormulonErrorCode::kParserUnexpectedToken).error();
  std::string out = cap.Read();
  EXPECT_TRUE(Contains(out, "\"level\":\"error\"")) << out;
  EXPECT_TRUE(Contains(out, "\"code\":1000")) << out;
  EXPECT_TRUE(Contains(out, "\"code_name\":\"kParserUnexpectedToken\"")) << out;
}

TEST(StructuredLogTest, StringFieldEscaped) {
  StderrCapture cap;
  StructuredLog("test.escape").field("msg", std::string_view("has \"quotes\" and a \\ backslash")).info();
  std::string out = cap.Read();
  EXPECT_TRUE(Contains(out, "has \\\"quotes\\\" and a \\\\ backslash")) << out;
}

TEST(StructuredLogTest, DebugWarnAndMultipleLevels) {
  StderrCapture cap;
  StructuredLog("a.b").debug();
  StructuredLog("c.d").warn();
  std::string out = cap.Read();
  EXPECT_TRUE(Contains(out, "\"level\":\"debug\"")) << out;
  EXPECT_TRUE(Contains(out, "\"level\":\"warn\"")) << out;
  EXPECT_TRUE(Contains(out, "\"event\":\"a.b\"")) << out;
  EXPECT_TRUE(Contains(out, "\"event\":\"c.d\"")) << out;
}

#else  // !FORMULON_HAVE_POSIX_PIPE

TEST(StructuredLogTest, SkippedOnNonPosixPlatforms) {
  GTEST_SKIP() << "stderr capture requires POSIX dup2/pipe";
}

#endif  // FORMULON_HAVE_POSIX_PIPE

}  // namespace
}  // namespace formulon
