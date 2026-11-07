//
// Process exit codes shared by the CLI subcommands. The values follow
// sysexits(3) where one fits; every other failure is the generic `1`.

#ifndef FORMULON_CLI_EXIT_CODES_H_
#define FORMULON_CLI_EXIT_CODES_H_

namespace formulon {
namespace cli {

/// Generic failure: the engine refused the request (including an empty or
/// corrupt input workbook).
inline constexpr int kExitFailure = 1;

/// Usage error. Mirrors sysexits(3) `EX_USAGE`.
inline constexpr int kExitUsage = 64;

/// The input text is malformed, such as a formula that does not parse. Mirrors
/// `EX_DATAERR`.
inline constexpr int kExitDataErr = 65;

/// An input file is missing, unreadable or a directory. Mirrors `EX_NOINPUT`.
inline constexpr int kExitNoInput = 66;

/// An output file cannot be created. Mirrors `EX_CANTCREAT`.
inline constexpr int kExitCantCreate = 73;

/// An I/O error while reading or writing a stream, including a closed pipe
/// on stdout. Mirrors `EX_IOERR`.
inline constexpr int kExitIoError = 74;

/// Maps an internal command status to the CLI's exit-code vocabulary. The
/// codes above pass through; every other non-zero value, such as an
/// `fm_status_t`, becomes `kExitFailure` because POSIX shells truncate a
/// status to eight bits and unrelated failures would collide.
constexpr int exit_code_for_status(int status) {
  switch (status) {
    case 0:
    case kExitUsage:
    case kExitDataErr:
    case kExitNoInput:
    case kExitCantCreate:
    case kExitIoError:
      return status;
    default:
      return kExitFailure;
  }
}

}  // namespace cli
}  // namespace formulon

#endif  // FORMULON_CLI_EXIT_CODES_H_
