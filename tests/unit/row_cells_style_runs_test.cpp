// Tests for `RowCells` style runs: blank columns carrying their own format,
// held beside the stored cells.
//
// The parameterised suite runs one case per row of a pairwise table over
// op {ensure, erase, erase_cols, shift_cols_from, erase_in, move_in_to,
// set_runs, clear_runs, remap_runs, demote} x run position {before,
// straddling_first, inside, straddling_last, after, spanning} x stored
// neighbour {none, left_adjacent, right_adjacent, inside_band} x adjacent run
// xf {equal, different, none}: 184/184 pairs in 61 rows (coverwise, seed
// 20261111). Each case applies the op to a `RowCells` and to a per-column
// reference model, then compares stored cells and runs and checks that no
// column is both stored and run and that runs are sorted, disjoint and merged.

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <vector>

#include "cell.h"
#include "gtest/gtest.h"
#include "row_cells.h"
#include "style_region.h"
#include "utils/a1_ref.h"
#include "value.h"

namespace formulon {
namespace {

constexpr std::uint32_t kPrimaryXf = 5U;
constexpr std::uint32_t kOtherXf = 7U;

// ---------------------------------------------------------------------------
// Reference model: one slot per column.

enum class Kind : std::uint8_t { kNone, kStored, kRun };

struct Slot {
  Kind kind = Kind::kNone;
  std::uint32_t xf = 0U;
  bool explicit_xf = false;
  int tag = 0;  // content marker (cached number); 0 = no content
};

using Model = std::vector<Slot>;

Model NewModel() {
  return Model(a1::kMaxCols);
}

std::uint64_t Clamp(std::uint64_t end) {
  return end < a1::kMaxCols ? end : a1::kMaxCols;
}

std::vector<StyleRun> ModelRuns(const Model& model) {
  std::vector<StyleRun> runs;
  for (std::uint32_t col = 0; col < model.size(); ++col) {
    const Slot& slot = model[col];
    if (slot.kind != Kind::kRun) {
      continue;
    }
    if (!runs.empty() && runs.back().last + 1U == col && runs.back().xf == slot.xf) {
      runs.back().last = static_cast<std::uint16_t>(col);
    } else {
      runs.push_back(StyleRun{static_cast<std::uint16_t>(col), static_cast<std::uint16_t>(col), slot.xf});
    }
  }
  return runs;
}

std::size_t Gain(std::size_t before, std::size_t after) {
  return after > before ? after - before : 0U;
}

void ModelEnsure(Model& model, std::uint32_t col) {
  Slot& slot = model[col];
  if (slot.kind == Kind::kRun) {
    slot.kind = Kind::kStored;
    slot.explicit_xf = true;
  } else if (slot.kind == Kind::kNone) {
    slot = Slot{Kind::kStored, 0U, false, 0};
  }
}

void ModelEraseIn(Model& model, std::uint64_t first, std::uint64_t end) {
  for (std::uint64_t col = first; col < Clamp(end); ++col) {
    model[col] = Slot{};
  }
}

void ModelEraseCols(Model& model, std::uint32_t col, std::uint32_t count) {
  model.erase(model.begin() + col, model.begin() + static_cast<std::ptrdiff_t>(Clamp(std::uint64_t{col} + count)));
  model.resize(a1::kMaxCols);
}

void ModelShift(Model& model, std::uint32_t col, std::uint32_t count) {
  model.insert(model.begin() + col, count, Slot{});
  model.resize(a1::kMaxCols);
}

void ModelSetRuns(Model& model, std::uint64_t first, std::uint64_t end, std::uint32_t xf) {
  for (std::uint64_t col = first; col < Clamp(end); ++col) {
    if (model[col].kind != Kind::kStored) {
      model[col] = Slot{Kind::kRun, xf, true, 0};
    }
  }
}

void ModelClearRuns(Model& model, std::uint64_t first, std::uint64_t end) {
  for (std::uint64_t col = first; col < Clamp(end); ++col) {
    if (model[col].kind == Kind::kRun) {
      model[col] = Slot{};
    }
  }
}

void ModelRemapRuns(Model& model, std::uint64_t first, std::uint64_t end, const XfMap& map) {
  for (std::uint64_t col = first; col < Clamp(end); ++col) {
    if (model[col].kind == Kind::kRun) {
      model[col].xf = map(model[col].xf);
    }
  }
}

void ModelMove(Model& source, std::uint64_t first, std::uint64_t end, Model& dest) {
  for (std::uint64_t col = first; col < Clamp(end); ++col) {
    dest[col] = source[col];
    source[col] = Slot{};
  }
}

std::optional<std::uint32_t> ModelDemote(Model& model, std::uint32_t col) {
  Slot& slot = model[col];
  if (slot.kind != Kind::kStored || slot.tag != 0) {
    return std::nullopt;
  }
  std::optional<std::uint32_t> owned;
  if (slot.xf != 0U || slot.explicit_xf) {
    owned = slot.xf;
  }
  slot = Slot{};
  return owned;
}

// ---------------------------------------------------------------------------
// Fixture writers and checks.

void Store(RowCells& row, Model& model, std::uint32_t col, int tag, std::uint32_t xf, bool explicit_xf) {
  Cell& cell = row.ensure(col);
  if (tag != 0) {
    cell.cached_value = Value::number(tag);
  }
  cell.xf_index = xf;
  cell.has_explicit_xf = explicit_xf;
  model[col] = Slot{Kind::kStored, xf, explicit_xf, tag};
}

void AddRun(RowCells& row, Model& model, std::uint32_t first, std::uint32_t last, std::uint32_t xf) {
  row.set_runs(first, std::uint64_t{last} + 1U, xf);
  ModelSetRuns(model, first, std::uint64_t{last} + 1U, xf);
}

int TagOf(const Cell& cell) {
  return cell.cached_value.is_number() ? static_cast<int>(cell.cached_value.as_number()) : 0;
}

// I1 and I2, read through the public surface only.
void ExpectInvariants(const RowCells& row) {
  const std::vector<StyleRun>& runs = row.runs();
  for (std::size_t i = 0; i < runs.size(); ++i) {
    EXPECT_LE(runs[i].first, runs[i].last) << "run " << i;
    EXPECT_LT(runs[i].last, a1::kMaxCols) << "run " << i;
    if (i != 0U) {
      EXPECT_LT(runs[i - 1U].last, runs[i].first) << "runs " << i - 1U << "/" << i << " overlap or are unsorted";
      EXPECT_FALSE(runs[i - 1U].last + 1U == runs[i].first && runs[i - 1U].xf == runs[i].xf)
          << "runs " << i - 1U << "/" << i << " touch with equal xf";
    }
  }
  for (const RowCells::Entry entry : row) {
    for (const StyleRun& run : runs) {
      EXPECT_FALSE(run.first <= entry.col && entry.col <= run.last) << "column " << entry.col << " stored and run";
    }
  }
  EXPECT_EQ(row.has_any(), !row.empty() || !runs.empty());
}

void ExpectRuns(const RowCells& row, const std::vector<StyleRun>& expected) {
  const std::vector<StyleRun>& runs = row.runs();
  ASSERT_EQ(runs.size(), expected.size());
  for (std::size_t i = 0; i < runs.size(); ++i) {
    EXPECT_EQ(runs[i].first, expected[i].first) << "run " << i;
    EXPECT_EQ(runs[i].last, expected[i].last) << "run " << i;
    EXPECT_EQ(runs[i].xf, expected[i].xf) << "run " << i;
  }
}

void ExpectMatches(const RowCells& row, const Model& model) {
  ExpectInvariants(row);
  ExpectRuns(row, ModelRuns(model));
  std::vector<std::uint32_t> expected_cols;
  for (std::uint32_t col = 0; col < model.size(); ++col) {
    if (model[col].kind == Kind::kStored) {
      expected_cols.push_back(col);
    }
  }
  std::vector<std::uint32_t> cols;
  for (const RowCells::Entry entry : row) {
    cols.push_back(entry.col);
    const Slot& slot = model[entry.col];
    EXPECT_EQ(entry.cell.xf_index, slot.xf) << "column " << entry.col;
    EXPECT_EQ(entry.cell.has_explicit_xf, slot.explicit_xf) << "column " << entry.col;
    EXPECT_EQ(TagOf(entry.cell), slot.tag) << "column " << entry.col;
  }
  EXPECT_EQ(cols, expected_cols);
}

// ---------------------------------------------------------------------------
// Pairwise cases.

enum class Op {
  kEnsure,
  kErase,
  kEraseCols,
  kShiftColsFrom,
  kEraseIn,
  kMoveInTo,
  kSetRuns,
  kClearRuns,
  kRemapRuns,
  kDemote,
};
enum class Pos { kBefore, kStraddlingFirst, kInside, kStraddlingLast, kAfter, kSpanning };
enum class Nb { kNone, kLeftAdjacent, kRightAdjacent, kInsideBand };
enum class Adj { kEqual, kDifferent, kNone };

struct Case {
  Op op;
  Pos pos;
  Nb nb;
  Adj adj;
};

constexpr Case kCases[] = {
    {Op::kMoveInTo, Pos::kAfter, Nb::kLeftAdjacent, Adj::kDifferent},
    {Op::kEnsure, Pos::kStraddlingFirst, Nb::kInsideBand, Adj::kDifferent},
    {Op::kEraseIn, Pos::kStraddlingLast, Nb::kLeftAdjacent, Adj::kNone},
    {Op::kClearRuns, Pos::kStraddlingFirst, Nb::kRightAdjacent, Adj::kNone},
    {Op::kMoveInTo, Pos::kSpanning, Nb::kRightAdjacent, Adj::kEqual},
    {Op::kDemote, Pos::kInside, Nb::kInsideBand, Adj::kEqual},
    {Op::kShiftColsFrom, Pos::kAfter, Nb::kRightAdjacent, Adj::kEqual},
    {Op::kShiftColsFrom, Pos::kSpanning, Nb::kInsideBand, Adj::kNone},
    {Op::kErase, Pos::kBefore, Nb::kInsideBand, Adj::kDifferent},
    {Op::kEnsure, Pos::kSpanning, Nb::kNone, Adj::kDifferent},
    {Op::kMoveInTo, Pos::kStraddlingLast, Nb::kInsideBand, Adj::kNone},
    {Op::kShiftColsFrom, Pos::kBefore, Nb::kLeftAdjacent, Adj::kEqual},
    {Op::kShiftColsFrom, Pos::kStraddlingFirst, Nb::kNone, Adj::kEqual},
    {Op::kDemote, Pos::kAfter, Nb::kNone, Adj::kNone},
    {Op::kClearRuns, Pos::kStraddlingLast, Nb::kNone, Adj::kEqual},
    {Op::kRemapRuns, Pos::kAfter, Nb::kInsideBand, Adj::kEqual},
    {Op::kSetRuns, Pos::kBefore, Nb::kNone, Adj::kNone},
    {Op::kSetRuns, Pos::kSpanning, Nb::kLeftAdjacent, Adj::kDifferent},
    {Op::kEraseIn, Pos::kInside, Nb::kRightAdjacent, Adj::kDifferent},
    {Op::kRemapRuns, Pos::kInside, Nb::kLeftAdjacent, Adj::kNone},
    {Op::kEnsure, Pos::kBefore, Nb::kRightAdjacent, Adj::kNone},
    {Op::kErase, Pos::kStraddlingLast, Nb::kRightAdjacent, Adj::kNone},
    {Op::kEraseCols, Pos::kInside, Nb::kNone, Adj::kDifferent},
    {Op::kMoveInTo, Pos::kStraddlingFirst, Nb::kNone, Adj::kDifferent},
    {Op::kRemapRuns, Pos::kStraddlingLast, Nb::kRightAdjacent, Adj::kDifferent},
    {Op::kDemote, Pos::kStraddlingFirst, Nb::kRightAdjacent, Adj::kDifferent},
    {Op::kEraseIn, Pos::kSpanning, Nb::kNone, Adj::kEqual},
    {Op::kEraseCols, Pos::kStraddlingFirst, Nb::kLeftAdjacent, Adj::kNone},
    {Op::kSetRuns, Pos::kAfter, Nb::kInsideBand, Adj::kEqual},
    {Op::kShiftColsFrom, Pos::kStraddlingLast, Nb::kInsideBand, Adj::kDifferent},
    {Op::kEraseCols, Pos::kStraddlingLast, Nb::kRightAdjacent, Adj::kEqual},
    {Op::kClearRuns, Pos::kInside, Nb::kRightAdjacent, Adj::kDifferent},
    {Op::kEraseIn, Pos::kAfter, Nb::kInsideBand, Adj::kEqual},
    {Op::kErase, Pos::kSpanning, Nb::kLeftAdjacent, Adj::kDifferent},
    {Op::kDemote, Pos::kSpanning, Nb::kLeftAdjacent, Adj::kNone},
    {Op::kEraseCols, Pos::kSpanning, Nb::kInsideBand, Adj::kDifferent},
    {Op::kEnsure, Pos::kStraddlingLast, Nb::kLeftAdjacent, Adj::kEqual},
    {Op::kSetRuns, Pos::kStraddlingFirst, Nb::kRightAdjacent, Adj::kNone},
    {Op::kErase, Pos::kInside, Nb::kNone, Adj::kNone},
    {Op::kClearRuns, Pos::kSpanning, Nb::kInsideBand, Adj::kEqual},
    {Op::kClearRuns, Pos::kBefore, Nb::kLeftAdjacent, Adj::kNone},
    {Op::kEraseIn, Pos::kBefore, Nb::kInsideBand, Adj::kNone},
    {Op::kClearRuns, Pos::kAfter, Nb::kInsideBand, Adj::kDifferent},
    {Op::kErase, Pos::kStraddlingFirst, Nb::kLeftAdjacent, Adj::kEqual},
    {Op::kMoveInTo, Pos::kInside, Nb::kLeftAdjacent, Adj::kEqual},
    {Op::kRemapRuns, Pos::kStraddlingLast, Nb::kNone, Adj::kNone},
    {Op::kEraseIn, Pos::kStraddlingFirst, Nb::kInsideBand, Adj::kNone},
    {Op::kErase, Pos::kAfter, Nb::kInsideBand, Adj::kNone},
    {Op::kEraseCols, Pos::kAfter, Nb::kNone, Adj::kEqual},
    {Op::kEnsure, Pos::kAfter, Nb::kInsideBand, Adj::kDifferent},
    {Op::kRemapRuns, Pos::kSpanning, Nb::kNone, Adj::kNone},
    {Op::kSetRuns, Pos::kInside, Nb::kInsideBand, Adj::kEqual},
    {Op::kSetRuns, Pos::kStraddlingLast, Nb::kInsideBand, Adj::kEqual},
    {Op::kRemapRuns, Pos::kBefore, Nb::kInsideBand, Adj::kDifferent},
    {Op::kRemapRuns, Pos::kStraddlingFirst, Nb::kNone, Adj::kNone},
    {Op::kEraseCols, Pos::kBefore, Nb::kLeftAdjacent, Adj::kNone},
    {Op::kShiftColsFrom, Pos::kInside, Nb::kLeftAdjacent, Adj::kEqual},
    {Op::kDemote, Pos::kStraddlingLast, Nb::kRightAdjacent, Adj::kEqual},
    {Op::kDemote, Pos::kBefore, Nb::kRightAdjacent, Adj::kDifferent},
    {Op::kEnsure, Pos::kInside, Nb::kRightAdjacent, Adj::kDifferent},
    {Op::kMoveInTo, Pos::kBefore, Nb::kLeftAdjacent, Adj::kDifferent},
};

const char* OpName(Op op) {
  switch (op) {
    case Op::kEnsure:
      return "ensure";
    case Op::kErase:
      return "erase";
    case Op::kEraseCols:
      return "erase_cols";
    case Op::kShiftColsFrom:
      return "shift_cols_from";
    case Op::kEraseIn:
      return "erase_in";
    case Op::kMoveInTo:
      return "move_in_to";
    case Op::kSetRuns:
      return "set_runs";
    case Op::kClearRuns:
      return "clear_runs";
    case Op::kRemapRuns:
      return "remap_runs";
    case Op::kDemote:
      return "demote";
  }
  return "";
}

const char* PosName(Pos pos) {
  switch (pos) {
    case Pos::kBefore:
      return "before";
    case Pos::kStraddlingFirst:
      return "straddling_first";
    case Pos::kInside:
      return "inside";
    case Pos::kStraddlingLast:
      return "straddling_last";
    case Pos::kAfter:
      return "after";
    case Pos::kSpanning:
      return "spanning";
  }
  return "";
}

const char* NbName(Nb nb) {
  switch (nb) {
    case Nb::kNone:
      return "none";
    case Nb::kLeftAdjacent:
      return "left_adjacent";
    case Nb::kRightAdjacent:
      return "right_adjacent";
    case Nb::kInsideBand:
      return "inside_band";
  }
  return "";
}

const char* AdjName(Adj adj) {
  switch (adj) {
    case Adj::kEqual:
      return "equal";
    case Adj::kDifferent:
      return "different";
    case Adj::kNone:
      return "none";
  }
  return "";
}

// The op band `[lo, hi)`: one column for the single-column ops, four for the
// column inserts and deletes, ten for the band ops.
struct Band {
  std::uint32_t lo;
  std::uint32_t hi;
};

Band BandOf(Op op) {
  switch (op) {
    case Op::kEnsure:
    case Op::kErase:
    case Op::kDemote:
      return {25U, 26U};
    case Op::kEraseCols:
    case Op::kShiftColsFrom:
      return {25U, 29U};
    default:
      return {20U, 30U};
  }
}

// The primary run, inclusive, placed against the band.
StyleRun PrimaryRun(Pos pos, Band band) {
  const std::uint32_t lo = band.lo;
  const std::uint32_t hi = band.hi;
  const bool wide = hi - lo > 1U;
  std::uint32_t first = 0;
  std::uint32_t last = 0;
  switch (pos) {
    case Pos::kBefore:
      first = lo - 6U;
      last = lo - 3U;
      break;
    case Pos::kStraddlingFirst:
      first = lo - 3U;
      last = wide ? lo + 1U : lo;
      break;
    case Pos::kInside:
      first = wide ? lo + 1U : lo;
      last = wide ? hi - 2U : lo;
      break;
    case Pos::kStraddlingLast:
      first = wide ? hi - 2U : lo;
      last = hi + 2U;
      break;
    case Pos::kAfter:
      first = hi + 2U;
      last = hi + 5U;
      break;
    case Pos::kSpanning:
      first = lo - 2U;
      last = hi + 1U;
      break;
  }
  return StyleRun{static_cast<std::uint16_t>(first), static_cast<std::uint16_t>(last), kPrimaryXf};
}

class RowCellsStyleRunsCases : public ::testing::TestWithParam<Case> {};

TEST_P(RowCellsStyleRunsCases, OpMatchesModel) {
  const Case& c = GetParam();
  const Band band = BandOf(c.op);
  const std::uint32_t width = band.hi - band.lo;
  const StyleRun primary = PrimaryRun(c.pos, band);

  RowCells row;
  Model model = NewModel();

  // Stored cells first, so the runs written next skip their columns.
  if (c.op == Op::kDemote) {
    Store(row, model, band.lo, 0, kPrimaryXf, true);
  }
  switch (c.nb) {
    case Nb::kNone:
      break;
    case Nb::kLeftAdjacent:
      Store(row, model, primary.first - 1U, 1, 0U, false);
      break;
    case Nb::kRightAdjacent:
      Store(row, model, primary.last + 1U, 1, 0U, false);
      break;
    case Nb::kInsideBand: {
      const std::uint32_t col = band.lo + width / 2U;
      if (model[col].kind == Kind::kNone) {
        Store(row, model, col, 1, 0U, false);
      }
      break;
    }
  }
  AddRun(row, model, primary.first, primary.last, primary.xf);
  if (c.adj != Adj::kNone) {
    AddRun(row, model, primary.last + 2U, primary.last + 4U, c.adj == Adj::kEqual ? kPrimaryXf : kOtherXf);
  }
  ExpectMatches(row, model);

  const std::size_t runs_before = ModelRuns(model).size();
  switch (c.op) {
    case Op::kEnsure: {
      const Cell& cell = row.ensure(band.lo);
      ModelEnsure(model, band.lo);
      EXPECT_EQ(cell.xf_index, model[band.lo].xf);
      EXPECT_EQ(cell.has_explicit_xf, model[band.lo].explicit_xf);
      break;
    }
    case Op::kErase:
      row.erase(band.lo);
      ModelEraseIn(model, band.lo, band.hi);
      break;
    case Op::kEraseCols:
      row.erase_cols(band.lo, width);
      ModelEraseCols(model, band.lo, width);
      break;
    case Op::kShiftColsFrom:
      row.shift_cols_from(band.lo, width);
      ModelShift(model, band.lo, width);
      break;
    case Op::kEraseIn:
      row.erase_in(band.lo, band.hi);
      ModelEraseIn(model, band.lo, band.hi);
      break;
    case Op::kMoveInTo: {
      RowCells dest;
      Model dest_model = NewModel();
      row.move_in_to(band.lo, band.hi, dest);
      ModelMove(model, band.lo, band.hi, dest_model);
      ExpectMatches(dest, dest_model);
      break;
    }
    case Op::kSetRuns: {
      const std::size_t added = row.set_runs(band.lo, band.hi, kOtherXf);
      ModelSetRuns(model, band.lo, band.hi, kOtherXf);
      EXPECT_EQ(added, Gain(runs_before, ModelRuns(model).size()));
      break;
    }
    case Op::kClearRuns:
      row.clear_runs(band.lo, band.hi);
      ModelClearRuns(model, band.lo, band.hi);
      break;
    case Op::kRemapRuns: {
      const XfRemap pairs[] = {{kPrimaryXf, kOtherXf}};
      const XfMap map{pairs, 1U, 0U, false};
      const std::size_t added = row.remap_runs(band.lo, band.hi, map);
      ModelRemapRuns(model, band.lo, band.hi, map);
      EXPECT_EQ(added, Gain(runs_before, ModelRuns(model).size()));
      break;
    }
    case Op::kDemote: {
      const std::optional<std::uint32_t> owned = row.demote(band.lo);
      const std::optional<std::uint32_t> expected = ModelDemote(model, band.lo);
      ASSERT_EQ(owned, expected);
      EXPECT_EQ(row.find(band.lo), nullptr);
      if (owned.has_value()) {
        // The caller writes an owned format back as a run.
        row.set_runs(band.lo, band.hi, *owned);
        ModelSetRuns(model, band.lo, band.hi, *owned);
      }
      break;
    }
  }
  ExpectMatches(row, model);
}

std::string CaseName(const ::testing::TestParamInfo<Case>& info) {
  std::string name = std::to_string(info.index);
  for (const char* part :
       {OpName(info.param.op), PosName(info.param.pos), NbName(info.param.nb), AdjName(info.param.adj)}) {
    name += '_';
    name += part;
  }
  return name;
}

INSTANTIATE_TEST_SUITE_P(Pairwise, RowCellsStyleRunsCases, ::testing::ValuesIn(kCases), CaseName);

TEST(RowCellsStyleRunsTest, PairwiseTableHasEveryRow) {
  EXPECT_EQ(sizeof(kCases) / sizeof(kCases[0]), 61U);
}

// ---------------------------------------------------------------------------
// Fixed tables.

TEST(RowCellsStyleRunsTest, MoveInToDestinationShapes) {
  enum class Dest { kEmpty, kRunsInBand, kRunsStraddlingBothEdges, kStoredInBand };
  struct Shape {
    Dest dest;
    const char* name;
    std::vector<StyleRun> expected_runs;
    std::vector<std::uint32_t> expected_cols;
  };
  const Shape shapes[] = {
      {Dest::kEmpty, "empty", {{20, 24, kPrimaryXf}, {26, 29, kPrimaryXf}}, {25U}},
      {Dest::kRunsInBand, "runs in band", {{20, 24, kPrimaryXf}, {26, 29, kPrimaryXf}}, {25U}},
      {Dest::kRunsStraddlingBothEdges,
       "runs straddling both edges",
       {{15, 24, kPrimaryXf}, {26, 35, kPrimaryXf}},
       {25U}},
      {Dest::kStoredInBand, "stored cells in band", {{20, 24, kPrimaryXf}, {26, 29, kPrimaryXf}}, {25U}},
  };
  for (const Shape& shape : shapes) {
    SCOPED_TRACE(shape.name);
    RowCells source;
    Model source_model = NewModel();
    Store(source, source_model, 25U, 1, 0U, false);
    AddRun(source, source_model, 18U, 32U, kPrimaryXf);

    RowCells dest;
    Model dest_model = NewModel();
    switch (shape.dest) {
      case Dest::kEmpty:
        break;
      case Dest::kRunsInBand:
        AddRun(dest, dest_model, 22U, 27U, kOtherXf);
        break;
      case Dest::kRunsStraddlingBothEdges:
        AddRun(dest, dest_model, 15U, 22U, kPrimaryXf);
        AddRun(dest, dest_model, 27U, 35U, kPrimaryXf);
        break;
      case Dest::kStoredInBand:
        Store(dest, dest_model, 21U, 2, 3U, true);
        Store(dest, dest_model, 25U, 3, 0U, false);
        break;
    }

    source.move_in_to(20U, 30U, dest);
    ModelMove(source_model, 20U, 30U, dest_model);
    ExpectMatches(source, source_model);
    ExpectMatches(dest, dest_model);
    ExpectRuns(source, {{18, 19, kPrimaryXf}, {30, 32, kPrimaryXf}});
    EXPECT_TRUE(source.empty());
    ExpectRuns(dest, shape.expected_runs);
    std::vector<std::uint32_t> cols;
    for (const RowCells::Entry entry : dest) {
      cols.push_back(entry.col);
      EXPECT_EQ(TagOf(entry.cell), 1);
    }
    EXPECT_EQ(cols, shape.expected_cols);
  }
}

TEST(RowCellsStyleRunsTest, RemapRunsMapShapes) {
  const XfRemap no_match[] = {{3U, 4U}};
  const XfRemap five_to_six[] = {{5U, 6U}};
  const XfRemap six_to_five[] = {{6U, 5U}};
  struct Shape {
    const char* name;
    bool two_runs;  // [10,14] xf 5 + [15,20] xf 6, else [10,30] xf 5
    std::uint32_t first;
    std::uint64_t end;
    XfMap map;
    std::vector<StyleRun> expected;
    std::size_t added;
  };
  const Shape shapes[] = {
      {"identity", false, 10U, 31U, XfMap{no_match, 1U, 0U, false}, {{10, 30, 5U}}, 0U},
      {"identity, empty table", false, 10U, 31U, XfMap{}, {{10, 30, 5U}}, 0U},
      {"splits one run",
       false,
       15U,
       20U,
       XfMap{five_to_six, 1U, 0U, false},
       {{10, 14, 5U}, {15, 19, 6U}, {20, 30, 5U}},
       2U},
      {"makes two neighbours equal", true, 15U, 21U, XfMap{six_to_five, 1U, 0U, false}, {{10, 20, 5U}}, 0U},
      {"maps to 0", false, 10U, 31U, XfMap{nullptr, 0U, 0U, true}, {{10, 30, 0U}}, 0U},
  };
  for (const Shape& shape : shapes) {
    SCOPED_TRACE(shape.name);
    RowCells row;
    if (shape.two_runs) {
      row.set_runs(10U, 15U, 5U);
      row.set_runs(15U, 21U, 6U);
    } else {
      row.set_runs(10U, 31U, 5U);
    }
    EXPECT_EQ(row.remap_runs(shape.first, shape.end, shape.map), shape.added);
    ExpectInvariants(row);
    ExpectRuns(row, shape.expected);
  }
}

TEST(RowCellsStyleRunsTest, XfMapForms) {
  const XfRemap pairs[] = {{1U, 2U}, {4U, 9U}};
  const XfMap table{pairs, 2U, 0U, false};
  EXPECT_EQ(table(1U), 2U);
  EXPECT_EQ(table(4U), 9U);
  EXPECT_EQ(table(3U), 3U);
  EXPECT_EQ(table(10U), 10U);
  const XfMap constant{pairs, 2U, 8U, true};
  EXPECT_EQ(constant(1U), 8U);
  EXPECT_EQ(constant(3U), 8U);
}

TEST(RowCellsStyleRunsTest, EmptySeesStoredCellsOnlyAndHasAnySeesRuns) {
  RowCells row;
  EXPECT_TRUE(row.empty());
  EXPECT_FALSE(row.has_any());
  EXPECT_TRUE(row.runs().empty());

  row.set_runs(3U, 6U, 4U);
  EXPECT_TRUE(row.empty());
  EXPECT_TRUE(row.has_any());
  EXPECT_EQ(row.stored_count(), 0U);
  EXPECT_TRUE(row.begin() == row.end());

  row.clear_runs(0U, a1::kMaxCols);
  EXPECT_FALSE(row.has_any());
  EXPECT_TRUE(row.runs().empty());

  row.ensure(9U);
  EXPECT_FALSE(row.empty());
  EXPECT_TRUE(row.has_any());
}

TEST(RowCellsStyleRunsTest, CapacityBytesCountsTheRunVector) {
  RowCells row;
  row.ensure(3U);
  const std::size_t base = row.capacity_bytes();
  row.set_runs(10U, 20U, 5U);
  EXPECT_GE(row.capacity_bytes(), base + sizeof(std::vector<StyleRun>) + sizeof(StyleRun));
  row.set_runs(25U, 30U, 6U);
  EXPECT_GE(row.capacity_bytes(), base + sizeof(std::vector<StyleRun>) + 2U * sizeof(StyleRun));
  // The vector is released with its last run.
  row.clear_runs(0U, a1::kMaxCols);
  EXPECT_EQ(row.capacity_bytes(), base);
}

TEST(RowCellsStyleRunsTest, ShiftTruncatesAtTheLastColumn) {
  constexpr std::uint32_t kLast = a1::kMaxCols - 1U;
  RowCells row;
  // An end past the grid is clamped.
  EXPECT_EQ(row.set_runs(kLast - 3U, std::uint64_t{a1::kMaxCols} + 100U, 5U), 1U);
  ExpectRuns(row, {{static_cast<std::uint16_t>(kLast - 3U), static_cast<std::uint16_t>(kLast), 5U}});

  // The piece at and after the shift point is pushed partly past the edge.
  row.shift_cols_from(kLast - 2U, 2U);
  ExpectInvariants(row);
  ExpectRuns(row, {{static_cast<std::uint16_t>(kLast - 3U), static_cast<std::uint16_t>(kLast - 3U), 5U},
                   {static_cast<std::uint16_t>(kLast), static_cast<std::uint16_t>(kLast), 5U}});

  // A run pushed fully past is dropped, and the vector with it.
  row.shift_cols_from(kLast - 3U, 4U);
  ExpectInvariants(row);
  EXPECT_TRUE(row.runs().empty());
  EXPECT_FALSE(row.has_any());

  // A stored cell and a run pushed past together.
  RowCells mixed;
  mixed.ensure(kLast - 1U);
  mixed.set_runs(kLast, a1::kMaxCols, 5U);
  mixed.shift_cols_from(kLast - 1U, 1U);
  ExpectInvariants(mixed);
  EXPECT_TRUE(mixed.runs().empty());
  ASSERT_NE(mixed.find(kLast), nullptr);
  EXPECT_EQ(mixed.stored_count(), 1U);
}

TEST(RowCellsStyleRunsTest, DemoteReturnsAnOwnedFormatAndErases) {
  RowCells row;
  Cell& styled = row.ensure(2U);
  styled.xf_index = 6U;
  row.ensure(3U);  // unstyled blank
  Cell& explicit_default = row.ensure(4U);
  explicit_default.has_explicit_xf = true;
  Cell& content = row.ensure(5U);
  content.cached_value = Value::number(1.0);
  content.xf_index = 6U;
  Cell& formula = row.ensure(6U);
  formula.formula_text = "=1";

  EXPECT_EQ(row.demote(2U), std::optional<std::uint32_t>(6U));
  EXPECT_EQ(row.find(2U), nullptr);
  EXPECT_EQ(row.demote(3U), std::nullopt);
  EXPECT_EQ(row.find(3U), nullptr);
  EXPECT_EQ(row.demote(4U), std::optional<std::uint32_t>(0U));
  EXPECT_EQ(row.find(4U), nullptr);

  // Content keeps the cell; nothing stored is a no-op.
  EXPECT_EQ(row.demote(5U), std::nullopt);
  EXPECT_NE(row.find(5U), nullptr);
  EXPECT_EQ(row.demote(6U), std::nullopt);
  EXPECT_NE(row.find(6U), nullptr);
  EXPECT_EQ(row.demote(7U), std::nullopt);
  EXPECT_EQ(row.stored_count(), 2U);
  EXPECT_TRUE(row.runs().empty());
}

TEST(RowCellsStyleRunsTest, RunsAddedReturnValues) {
  RowCells row;
  EXPECT_EQ(row.set_runs(10U, 20U, 5U), 1U);
  // A different xf inside one run splits it: one run becomes three.
  EXPECT_EQ(row.set_runs(12U, 15U, 7U), 2U);
  ExpectRuns(row, {{10, 11, 5U}, {12, 14, 7U}, {15, 19, 5U}});
  // Rewriting the middle back merges the three into one.
  EXPECT_EQ(row.set_runs(12U, 15U, 5U), 0U);
  ExpectRuns(row, {{10, 19, 5U}});
  // An equal xf touching the run extends it.
  EXPECT_EQ(row.set_runs(20U, 25U, 5U), 0U);
  ExpectRuns(row, {{10, 24, 5U}});
  // Stored columns split the coverage.
  RowCells gaps;
  gaps.ensure(12U);
  gaps.ensure(15U);
  EXPECT_EQ(gaps.set_runs(10U, 20U, 5U), 3U);
  ExpectInvariants(gaps);
  ExpectRuns(gaps, {{10, 11, 5U}, {13, 14, 5U}, {16, 19, 5U}});
  // A band holding only stored columns adds nothing.
  RowCells full;
  full.ensure(4U);
  EXPECT_EQ(full.set_runs(4U, 5U, 5U), 0U);
  EXPECT_TRUE(full.runs().empty());
  // Empty band.
  EXPECT_EQ(row.set_runs(30U, 30U, 5U), 0U);
}

TEST(RowCellsStyleRunsTest, EnsureInsideARunSplitsItAndKeepsTheFormat) {
  RowCells row;
  row.set_runs(10U, 20U, 5U);
  Cell& cell = row.ensure(15U);
  EXPECT_EQ(cell.xf_index, 5U);
  EXPECT_TRUE(cell.has_explicit_xf);
  ExpectInvariants(row);
  ExpectRuns(row, {{10, 14, 5U}, {16, 19, 5U}});
  // At a run edge only the edge column leaves.
  row.ensure(10U);
  ExpectRuns(row, {{11, 14, 5U}, {16, 19, 5U}});
  // A single-column run disappears entirely.
  RowCells single;
  single.set_runs(3U, 4U, 9U);
  EXPECT_EQ(single.ensure(3U).xf_index, 9U);
  EXPECT_TRUE(single.runs().empty());
  // Outside every run a default cell is made.
  const Cell& plain = row.ensure(30U);
  EXPECT_EQ(plain.xf_index, 0U);
  EXPECT_FALSE(plain.has_explicit_xf);
}

TEST(RowCellsStyleRunsTest, StyleQueries) {
  RowCells row;
  row.set_runs(10U, 15U, 5U);
  row.set_runs(20U, 25U, 7U);
  row.ensure(17U);
  Cell& explicit_cell = row.ensure(18U);
  explicit_cell.has_explicit_xf = true;
  explicit_cell.xf_index = 0U;

  const auto [lo, hi] = row.runs_in(12U, 22U);
  ASSERT_EQ(hi - lo, 2);
  EXPECT_EQ(lo[0].first, 10U);  // unclipped
  EXPECT_EQ(lo[1].last, 24U);
  const auto [gap_lo, gap_hi] = row.runs_in(15U, 20U);
  EXPECT_EQ(gap_lo, gap_hi);
  EXPECT_EQ(row.run_count_in(12U, 22U), 5U);
  EXPECT_EQ(row.run_count_in(0U, a1::kMaxCols), 10U);
  EXPECT_EQ(row.run_count_in(15U, 20U), 0U);

  EXPECT_FALSE(row.any_in(15U, 17U));
  EXPECT_TRUE(row.any_in(15U, 18U));
  EXPECT_TRUE(row.any_in(14U, 15U));
  EXPECT_FALSE(row.any_in(25U, a1::kMaxCols));

  // A stored cell owns its style even with xf 0 and no explicit flag.
  const std::optional<StyleSlot> stored = row.find_style(17U);
  ASSERT_TRUE(stored.has_value());
  EXPECT_EQ(stored->xf, 0U);
  EXPECT_FALSE(stored->explicit_xf);
  const std::optional<StyleSlot> explicit_default = row.find_style(18U);
  ASSERT_TRUE(explicit_default.has_value());
  EXPECT_TRUE(explicit_default->explicit_xf);
  const std::optional<StyleSlot> run = row.find_style(11U);
  ASSERT_TRUE(run.has_value());
  EXPECT_EQ(run->xf, 5U);
  EXPECT_TRUE(run->explicit_xf);
  EXPECT_FALSE(row.find_style(16U).has_value());

  RowCells none;
  EXPECT_FALSE(none.find_style(0U).has_value());
  EXPECT_EQ(none.run_count_in(0U, a1::kMaxCols), 0U);
  const auto [none_lo, none_hi] = none.runs_in(0U, a1::kMaxCols);
  EXPECT_EQ(none_lo, none_hi);
}

TEST(RowCellsStyleRunsTest, EraseColsJoinsRunsAcrossTheBand) {
  RowCells row;
  row.set_runs(5U, 10U, 5U);
  row.set_runs(13U, 20U, 5U);
  row.erase_cols(10U, 3U);
  ExpectInvariants(row);
  ExpectRuns(row, {{5, 16, 5U}});
}

}  // namespace
}  // namespace formulon
