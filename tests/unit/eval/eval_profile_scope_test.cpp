// Regression tests for carrying a workbook's Excel profile into evaluation
// contexts used by conditional formatting and data validation.

#include "eval/eval_profile_scope.h"

#include <array>
#include <chrono>
#include <condition_variable>
#include <cstddef>
#include <cstdint>
#include <mutex>
#include <string_view>
#include <type_traits>
#include <utility>

#include "cf/cf_evaluator.h"
#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/parser.h"
#include "utils/arena.h"
#include "utils/thread_launch.h"
#include "workbook.h"

namespace formulon::eval {
namespace {

struct ProfileCase {
  ExcelProfile profile;
  const char* id;
};

constexpr std::array<ProfileCase, 4> kProfiles = {
    ProfileCase{mac_365_ja_jp_profile(), "mac-365-ja_JP"},
    ProfileCase{win_365_ja_jp_profile(), "win-365-ja_JP"},
    ProfileCase{mac_365_en_us_profile(), "mac-365-en_US"},
    ProfileCase{win_365_en_us_profile(), "win-365-en_US"},
};

class TwoPartyGate {
 public:
  bool arrive_and_wait() {
    std::unique_lock<std::mutex> lock(mutex_);
    ++arrived_;
    if (arrived_ == 2U) {
      open_ = true;
      condition_.notify_all();
    } else if (!condition_.wait_for(lock, std::chrono::seconds(2), [this] { return open_; })) {
      timed_out_ = true;
      open_ = true;
      condition_.notify_all();
    }
    return open_ && !timed_out_;
  }

  bool timed_out() const {
    std::lock_guard<std::mutex> lock(mutex_);
    return timed_out_;
  }

 private:
  mutable std::mutex mutex_;
  std::condition_variable condition_;
  std::size_t arrived_ = 0U;
  bool open_ = false;
  bool timed_out_ = false;
};

struct ProfileGate {
  TwoPartyGate before_read;
  TwoPartyGate after_read;
};

thread_local ProfileGate* g_profile_gate = nullptr;

Value CurrentProfileImpl(const Value* /*args*/, std::uint32_t /*arity*/, Arena& /*arena*/) {
  if (g_profile_gate != nullptr) {
    g_profile_gate->before_read.arrive_and_wait();
  }
  const Value result = Value::text(excel_profile_id(current_eval_profile()));
  if (g_profile_gate != nullptr) {
    g_profile_gate->after_read.arrive_and_wait();
  }
  return result;
}

Value ProfileErrorImpl(const Value* /*args*/, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::error(ErrorCode::Value);
}

Value EvaluateFormula(std::string_view formula, const FunctionRegistry& registry, ExcelProfile profile,
                      bool first_element_only) {
  Arena parse_arena;
  parser::AstNode* root = parser::parse_strict(formula, parse_arena);
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  Arena eval_arena;
  const EvalContext context = EvalContext().with_excel_profile(profile);
  if (first_element_only) {
    return evaluate_first_element(*root, eval_arena, registry, context);
  }
  return evaluate(*root, eval_arena, registry, context);
}

struct ThreadResult {
  std::string_view observed;
  std::string_view after_scope;
  bool completed = false;
  bool restored = false;
};

struct ThreadArgs {
  std::size_t index;
  ExcelProfile profile;
  bool first_element_only;
  const FunctionRegistry* registry;
  ProfileGate* gate;
  std::array<ThreadResult, 2>* results;
};

void EvaluateOnThread(void* raw) {
  const ThreadArgs& args = *static_cast<const ThreadArgs*>(raw);
  Arena parse_arena;
  parser::AstNode* root = parser::parse_strict("CURRENT_PROFILE()", parse_arena);
  if (root == nullptr) {
    return;
  }
  Arena eval_arena;
  const EvalContext context = EvalContext().with_excel_profile(args.profile);
  g_profile_gate = args.gate;
  const Value value = args.first_element_only ? evaluate_first_element(*root, eval_arena, *args.registry, context)
                                              : evaluate(*root, eval_arena, *args.registry, context);
  const ExcelProfile after_scope = current_eval_profile();
  g_profile_gate = nullptr;
  (*args.results)[args.index].after_scope = excel_profile_id(after_scope);
  (*args.results)[args.index].restored = same_profile(after_scope, default_excel_profile());
  if (value.is_text()) {
    (*args.results)[args.index].observed = value.as_text();
    (*args.results)[args.index].completed = true;
  }
}

static_assert(!std::is_copy_constructible_v<ScopedEvalProfile>);
static_assert(!std::is_copy_assignable_v<ScopedEvalProfile>);
static_assert(!std::is_move_constructible_v<ScopedEvalProfile>);
static_assert(!std::is_move_assignable_v<ScopedEvalProfile>);

TEST(EvalProfilePropagation, ConditionalFormatHostContextUsesWorkbookProfile) {
  Workbook wb = Workbook::create();
  wb.set_excel_profile(mac_365_ja_jp_profile());
  Sheet& sheet = wb.sheet(0);

  // Mirror the context and host construction used by the C conditional-format
  // entry point: the host forwards one workbook-aware three-argument context
  // to every rule evaluation.
  Arena arena;
  EvalState state;
  EvalContext eval_ctx(wb, sheet, state);
  cf::CFHost host;
  host.arena = &arena;
  host.registry = &default_registry();
  host.eval_ctx = &eval_ctx;

  ASSERT_EQ(host.eval_ctx, &eval_ctx);
  ASSERT_EQ(host.eval_ctx->workbook(), &wb);
  ASSERT_EQ(host.eval_ctx->current_sheet(), &sheet);
  EXPECT_TRUE(same_profile(host.eval_ctx->excel_profile(), wb.excel_profile()));
  EXPECT_EQ(host.eval_ctx->excel_profile().host, ExcelHost::kMac365);
}

TEST(EvalProfilePropagation, ValidationThreeArgumentContextUsesWorkbookProfile) {
  Workbook wb = Workbook::create();
  wb.set_excel_profile(mac_365_ja_jp_profile());
  Sheet& sheet = wb.sheet(0);

  // The data-validation evaluator creates this exact workbook/current-sheet/
  // state context before evaluating a rule formula.
  EvalState state;
  EvalContext eval_ctx(wb, sheet, state);

  ASSERT_EQ(eval_ctx.workbook(), &wb);
  ASSERT_EQ(eval_ctx.current_sheet(), &sheet);
  EXPECT_TRUE(same_profile(eval_ctx.excel_profile(), wb.excel_profile()));
  EXPECT_EQ(eval_ctx.excel_profile().host, ExcelHost::kMac365);
}

TEST(EvalProfilePropagation, WorkbookOnlyContextUsesWorkbookProfile) {
  Workbook wb = Workbook::create();
  wb.set_excel_profile(mac_365_ja_jp_profile());
  Sheet& sheet = wb.sheet(0);

  const EvalContext eval_ctx = EvalContext::workbook_only(wb, sheet);

  ASSERT_EQ(eval_ctx.workbook(), &wb);
  ASSERT_EQ(eval_ctx.current_sheet(), &sheet);
  EXPECT_TRUE(same_profile(eval_ctx.excel_profile(), wb.excel_profile()));
  EXPECT_EQ(eval_ctx.excel_profile().host, ExcelHost::kMac365);
}

TEST(EvalProfileScope, DefaultsToTheRuntimeProfileOutsideEvaluation) {
  EXPECT_TRUE(same_profile(current_eval_profile(), default_excel_profile()));
}

TEST(EvalProfileScope, NestedScopesRestoreThePreviousProfile) {
  const ExcelProfile before = current_eval_profile();
  {
    ScopedEvalProfile outer(mac_365_ja_jp_profile());
    EXPECT_TRUE(same_profile(current_eval_profile(), mac_365_ja_jp_profile()));
    {
      ScopedEvalProfile inner(win_365_en_us_profile());
      EXPECT_TRUE(same_profile(current_eval_profile(), win_365_en_us_profile()));
    }
    EXPECT_TRUE(same_profile(current_eval_profile(), mac_365_ja_jp_profile()));
  }
  EXPECT_TRUE(same_profile(current_eval_profile(), before));
}

TEST(EvalProfileScope, CustomFunctionSeesContextProfileThroughBothEvaluationEntrypoints) {
  FunctionRegistry registry;
  ASSERT_TRUE(registry.register_function(FunctionDef{"CURRENT_PROFILE", 0U, 0U, &CurrentProfileImpl}));

  for (const ProfileCase& profile : kProfiles) {
    const Value scalar = EvaluateFormula("CURRENT_PROFILE()", registry, profile.profile, false);
    ASSERT_TRUE(scalar.is_text()) << profile.id;
    EXPECT_EQ(scalar.as_text(), profile.id);

    const Value first = EvaluateFormula("CURRENT_PROFILE()", registry, profile.profile, true);
    ASSERT_TRUE(first.is_text()) << profile.id;
    EXPECT_EQ(first.as_text(), profile.id);
    EXPECT_TRUE(same_profile(current_eval_profile(), default_excel_profile()));
  }
}

TEST(EvalProfileScope, EvaluationRestoresAnOuterProfileAfterSuccessAndError) {
  FunctionRegistry registry;
  ASSERT_TRUE(registry.register_function(FunctionDef{"CURRENT_PROFILE", 0U, 0U, &CurrentProfileImpl}));
  ASSERT_TRUE(registry.register_function(FunctionDef{"PROFILE_ERROR", 0U, 0U, &ProfileErrorImpl}));

  ScopedEvalProfile outer(mac_365_ja_jp_profile());
  const Value success = EvaluateFormula("CURRENT_PROFILE()", registry, win_365_en_us_profile(), false);
  ASSERT_TRUE(success.is_text());
  EXPECT_EQ(success.as_text(), "win-365-en_US");
  EXPECT_TRUE(same_profile(current_eval_profile(), mac_365_ja_jp_profile()));

  const Value error = EvaluateFormula("PROFILE_ERROR()", registry, mac_365_en_us_profile(), true);
  ASSERT_TRUE(error.is_error());
  EXPECT_EQ(error.as_error(), ErrorCode::Value);
  EXPECT_TRUE(same_profile(current_eval_profile(), mac_365_ja_jp_profile()));
}

TEST(EvalProfileScope, ConcurrentEvaluationsKeepProfilesIndependent) {
  FunctionRegistry registry;
  ASSERT_TRUE(registry.register_function(FunctionDef{"CURRENT_PROFILE", 0U, 0U, &CurrentProfileImpl}));

  ProfileGate gate;
  std::array<ThreadResult, 2> results{};
  ThreadArgs mac_args{0U, mac_365_en_us_profile(), false, &registry, &gate, &results};
  ThreadArgs win_args{1U, win_365_ja_jp_profile(), true, &registry, &gate, &results};
  ThreadStart mac_start{&EvaluateOnThread, &mac_args};
  ThreadStart win_start{&EvaluateOnThread, &win_args};
  auto mac_thread_or = launch_thread(mac_start);
  ASSERT_TRUE(static_cast<bool>(mac_thread_or)) << mac_thread_or.error().message;
  Thread mac_thread = std::move(mac_thread_or.value());
  auto win_thread_or = launch_thread(win_start);
  if (!static_cast<bool>(win_thread_or)) {
    mac_thread.join();
    FAIL() << win_thread_or.error().message;
  }
  Thread win_thread = std::move(win_thread_or.value());
  mac_thread.join();
  win_thread.join();

  EXPECT_FALSE(gate.before_read.timed_out());
  EXPECT_FALSE(gate.after_read.timed_out());
  ASSERT_TRUE(results[0].completed);
  ASSERT_TRUE(results[1].completed);
  EXPECT_TRUE(results[0].restored);
  EXPECT_TRUE(results[1].restored);
  EXPECT_EQ(results[0].observed, "mac-365-en_US");
  EXPECT_EQ(results[1].observed, "win-365-ja_JP");
  EXPECT_EQ(results[0].after_scope, "win-365-ja_JP");
  EXPECT_EQ(results[1].after_scope, "win-365-ja_JP");
}

TEST(EvalProfilePropagation, WorkbookBoundContextsUseEveryProfile) {
  for (const ProfileCase& profile : kProfiles) {
    Workbook wb = Workbook::create();
    wb.set_excel_profile(profile.profile);
    Sheet& sheet = wb.sheet(0);
    EvalState state;
    const EvalContext stateful_context(wb, sheet, state);
    const EvalContext workbook_only_context = EvalContext::workbook_only(wb, sheet);

    EXPECT_TRUE(same_profile(stateful_context.excel_profile(), profile.profile)) << profile.id;
    EXPECT_TRUE(same_profile(workbook_only_context.excel_profile(), profile.profile)) << profile.id;

    Arena arena;
    cf::CFHost host;
    host.arena = &arena;
    host.registry = &default_registry();
    host.eval_ctx = &stateful_context;
    ASSERT_EQ(host.eval_ctx, &stateful_context);
    EXPECT_TRUE(same_profile(host.eval_ctx->excel_profile(), profile.profile)) << profile.id;
  }
}

}  // namespace
}  // namespace formulon::eval
