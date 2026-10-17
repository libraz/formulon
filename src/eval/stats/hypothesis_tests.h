//
// Sample-statistics kernels behind T.TEST, F.TEST, CHISQ.TEST and the
// CONFIDENCE family, evaluated operation by operation (no fused multiply-add,
// subnormal products flushed) so results reproduce Excel's bit for bit.

#ifndef FORMULON_EVAL_STATS_HYPOTHESIS_TESTS_H_
#define FORMULON_EVAL_STATS_HYPOTHESIS_TESTS_H_

#include <cstddef>

namespace formulon {
namespace eval {
namespace stats {

/// Outcome of a sample test: `p` is meaningful only when `div0` is false.
struct TestOutcome {
  double p;
  bool div0;
};

/// T.TEST type 1 on `n` pairs; `tails` is 1 or 2.
TestOutcome t_test_paired_excel(const double* x, const double* y, std::size_t n, int tails) noexcept;

/// T.TEST type 2 (pooled) or 3 (Welch) on two independent samples.
TestOutcome t_test_two_sample_excel(const double* x, std::size_t nx, const double* y, std::size_t ny, int tails,
                                    int type) noexcept;

/// F.TEST: the two-tailed probability that the sample variances differ.
TestOutcome f_test_excel(const double* x, std::size_t nx, const double* y, std::size_t ny) noexcept;

/// One term of the chi-squared statistic added to `acc`: `((o - e)^2 / e) + acc`.
double chisq_accumulate_excel(double acc, double observed, double expected) noexcept;

/// CHISQ.TEST p-value: upper tail of the chi-squared law at `stat` with `df` degrees of freedom.
double chisq_test_p_excel(double stat, double df) noexcept;

/// CONFIDENCE.NORM: `(sd / sqrt(size)) * z`; `size` truncated, `size >= 1`.
double confidence_norm_excel(double alpha, double sd, double size) noexcept;

/// CONFIDENCE.T: `(sd / sqrt(size)) * T.INV.2T(alpha, size - 1)`.
double confidence_t_excel(double alpha, double sd, double size) noexcept;

}  // namespace stats
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_STATS_HYPOTHESIS_TESTS_H_
