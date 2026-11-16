// Empty strings in the comment, person, hyperlink, validation, defined-name, protection, table and format-reference
// setters (api-conventions §5 rule 30): where "" is not a legal value it is refused and the absent value (NULL)
// clears or derives; where the file holds one state for both they mean the same; where the file keeps "" it is a value.

#include <cstddef>
#include <cstdint>
#include <string>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);

struct Book {
  fm_workbook_t* handle = nullptr;
  Book() { EXPECT_EQ(fm_workbook_create_default(&handle), 0); }
  ~Book() { fm_workbook_destroy(handle); }
  Book(const Book&) = delete;
  Book& operator=(const Book&) = delete;
};

constexpr fm_cell_rect_t kCell = {0, 0, 0, 0};

// ---------------------------------------------------------------------------
// format_reference
// ---------------------------------------------------------------------------

TEST(TextFieldsEmptyStrings, FormatReferenceRefusesAnEmptySheetNameAndNullWritesNoQualifier) {
  Book wb;
  const fm_cell_rect_t rect = {1, 1, 2, 2};
  const char* text = nullptr;
  EXPECT_EQ(fm_workbook_format_reference(wb.handle, "", &rect, 0, &text), kInvalidArgument);
  ASSERT_EQ(fm_workbook_format_reference(wb.handle, nullptr, &rect, 0, &text), 0);
  EXPECT_STREQ(text, "B2:C3");
  ASSERT_EQ(fm_workbook_format_reference(wb.handle, "Sheet1", &rect, 0, &text), 0);
  EXPECT_STREQ(text, "Sheet1!B2:C3");
}

// ---------------------------------------------------------------------------
// Hyperlink
// ---------------------------------------------------------------------------

size_t HyperlinkCount(const fm_workbook_t* wb) {
  size_t count = 99;
  EXPECT_EQ(fm_sheet_get_hyperlink_count(wb, 0, &count), 0);
  return count;
}

TEST(TextFieldsEmptyStrings, HyperlinkRefusesAnEmptyTargetAndLocation) {
  Book wb;
  fm_hyperlink_t link = FM_STRUCT_INIT(fm_hyperlink_t);
  link.rect = kCell;
  link.target = "";
  EXPECT_EQ(fm_sheet_set_hyperlink(wb.handle, 0, &link), kInvalidArgument);
  EXPECT_EQ(fm_sheet_insert_hyperlink_at(wb.handle, 0, 0, &link), kInvalidArgument);

  link.target = "https://example.com/";
  link.location = "";
  EXPECT_EQ(fm_sheet_set_hyperlink(wb.handle, 0, &link), kInvalidArgument);
  EXPECT_EQ(fm_sheet_insert_hyperlink_at(wb.handle, 0, 0, &link), kInvalidArgument);

  link.target = nullptr;
  EXPECT_EQ(fm_sheet_set_hyperlink(wb.handle, 0, &link), kInvalidArgument);
  EXPECT_EQ(HyperlinkCount(wb.handle), 0U);
}

TEST(TextFieldsEmptyStrings, HyperlinkNullTargetOrLocationIsAbsent) {
  Book wb;
  fm_hyperlink_t link = FM_STRUCT_INIT(fm_hyperlink_t);
  link.rect = kCell;
  link.target = "https://example.com/";
  ASSERT_EQ(fm_sheet_set_hyperlink(wb.handle, 0, &link), 0) << fm_get_last_error_message();

  link = FM_STRUCT_INIT(fm_hyperlink_t);
  link.rect = fm_cell_rect_t{0, 1, 0, 1};
  link.location = "Sheet1!C3";
  ASSERT_EQ(fm_sheet_set_hyperlink(wb.handle, 0, &link), 0) << fm_get_last_error_message();

  int32_t present = 0;
  fm_hyperlink_t out = FM_STRUCT_INIT(fm_hyperlink_t);
  ASSERT_EQ(fm_sheet_get_hyperlink(wb.handle, 0, 0, 0, &present, &out), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(out.location, nullptr);
  ASSERT_EQ(fm_sheet_get_hyperlink(wb.handle, 0, 0, 1, &present, &out), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(out.target, nullptr);
  EXPECT_STREQ(out.location, "Sheet1!C3");
}

TEST(TextFieldsEmptyStrings, HyperlinkEmptyDisplayAndTooltipAreTheAbsentValue) {
  Book wb;
  fm_hyperlink_t empty = FM_STRUCT_INIT(fm_hyperlink_t);
  empty.rect = kCell;
  empty.target = "https://example.com/";
  empty.display = "";
  empty.tooltip = "";
  ASSERT_EQ(fm_sheet_set_hyperlink(wb.handle, 0, &empty), 0) << fm_get_last_error_message();
  fm_hyperlink_t omitted = FM_STRUCT_INIT(fm_hyperlink_t);
  omitted.rect = fm_cell_rect_t{0, 1, 0, 1};
  omitted.target = "https://example.com/";
  ASSERT_EQ(fm_sheet_set_hyperlink(wb.handle, 0, &omitted), 0) << fm_get_last_error_message();

  int32_t present = 0;
  fm_hyperlink_t a = FM_STRUCT_INIT(fm_hyperlink_t);
  fm_hyperlink_t b = FM_STRUCT_INIT(fm_hyperlink_t);
  ASSERT_EQ(fm_sheet_get_hyperlink(wb.handle, 0, 0, 0, &present, &a), 0);
  ASSERT_EQ(fm_sheet_get_hyperlink(wb.handle, 0, 0, 1, &present, &b), 0);
  EXPECT_EQ(a.tooltip, nullptr);
  EXPECT_EQ(b.tooltip, nullptr);
  EXPECT_EQ(a.display == nullptr, b.display == nullptr);
  if (a.display != nullptr && b.display != nullptr) {
    EXPECT_STREQ(a.display, b.display);
  }
}

// ---------------------------------------------------------------------------
// Data validation
// ---------------------------------------------------------------------------

fm_data_validation_t Rule(const fm_cell_rect_t* range) {
  fm_data_validation_t v = FM_STRUCT_INIT(fm_data_validation_t);
  v.ranges = range;
  v.range_count = 1;
  v.type = FM_VALIDATION_TYPE_WHOLE;
  v.formula1 = "1";
  v.formula2 = "5";
  return v;
}

size_t ValidationCount(const fm_workbook_t* wb) {
  size_t count = 99;
  EXPECT_EQ(fm_sheet_get_validation_count(wb, 0, &count), 0);
  return count;
}

TEST(TextFieldsEmptyStrings, ValidationAddRefusesAnEmptyFormula) {
  Book wb;
  const fm_cell_rect_t range = {0, 3, 2, 3};
  const char* id = nullptr;
  fm_data_validation_t v = Rule(&range);
  v.formula1 = "";
  EXPECT_EQ(fm_sheet_add_validation(wb.handle, 0, &v, &id), kInvalidArgument);
  v = Rule(&range);
  v.formula2 = "";
  EXPECT_EQ(fm_sheet_add_validation(wb.handle, 0, &v, &id), kInvalidArgument);
  EXPECT_EQ(ValidationCount(wb.handle), 0U);

  v = Rule(&range);
  v.formula2 = nullptr;
  ASSERT_EQ(fm_sheet_add_validation(wb.handle, 0, &v, &id), 0) << fm_get_last_error_message();
  fm_data_validation_t out = FM_STRUCT_INIT(fm_data_validation_t);
  ASSERT_EQ(fm_sheet_get_validation_at(wb.handle, 0, 0, &out), 0);
  EXPECT_NE(out.formula1, nullptr);
  EXPECT_EQ(out.formula2, nullptr);
}

TEST(TextFieldsEmptyStrings, ValidationUpdateRefusesAnEmptyFormulaAndClearRemovesIt) {
  Book wb;
  const fm_cell_rect_t range = {0, 3, 2, 3};
  const char* id_view = nullptr;
  fm_data_validation_t v = Rule(&range);
  ASSERT_EQ(fm_sheet_add_validation(wb.handle, 0, &v, &id_view), 0) << fm_get_last_error_message();
  const std::string id = id_view;

  fm_data_validation_patch_t patch = FM_STRUCT_INIT(fm_data_validation_patch_t);
  patch.formula1_op = FM_PATCH_OP_SET;
  patch.formula1 = "";
  EXPECT_EQ(fm_sheet_update_validation(wb.handle, 0, id.c_str(), &patch), kInvalidArgument);
  patch = FM_STRUCT_INIT(fm_data_validation_patch_t);
  patch.formula2_op = FM_PATCH_OP_SET;
  patch.formula2 = "";
  EXPECT_EQ(fm_sheet_update_validation(wb.handle, 0, id.c_str(), &patch), kInvalidArgument);

  fm_data_validation_t out = FM_STRUCT_INIT(fm_data_validation_t);
  ASSERT_EQ(fm_sheet_get_validation_at(wb.handle, 0, 0, &out), 0);
  EXPECT_NE(out.formula1, nullptr);
  EXPECT_NE(out.formula2, nullptr);

  patch = FM_STRUCT_INIT(fm_data_validation_patch_t);
  patch.formula2_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(fm_sheet_update_validation(wb.handle, 0, id.c_str(), &patch), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_validation_at(wb.handle, 0, 0, &out), 0);
  EXPECT_EQ(out.formula2, nullptr);
}

TEST(TextFieldsEmptyStrings, ValidationEmptyTitlesAndMessagesAreTheAbsentValue) {
  Book wb;
  const fm_cell_rect_t range = {0, 3, 2, 3};
  const char* id = nullptr;
  fm_data_validation_t v = Rule(&range);
  v.error_title = "";
  v.error_message = "";
  v.prompt_title = "";
  v.prompt_message = "";
  ASSERT_EQ(fm_sheet_add_validation(wb.handle, 0, &v, &id), 0) << fm_get_last_error_message();
  fm_data_validation_t out = FM_STRUCT_INIT(fm_data_validation_t);
  ASSERT_EQ(fm_sheet_get_validation_at(wb.handle, 0, 0, &out), 0);
  EXPECT_EQ(out.error_title, nullptr);
  EXPECT_EQ(out.error_message, nullptr);
  EXPECT_EQ(out.prompt_title, nullptr);
  EXPECT_EQ(out.prompt_message, nullptr);
}

// ---------------------------------------------------------------------------
// Defined name
// ---------------------------------------------------------------------------

TEST(TextFieldsEmptyStrings, DefinedNameEmptyCommentAndNullAreOneState) {
  Book wb;
  fm_defined_name_t d = FM_STRUCT_INIT(fm_defined_name_t);
  d.name = "WithEmpty";
  d.formula = "=Sheet1!$A$1";
  d.local_sheet_id = FM_DEFINED_NAME_SCOPE_WORKBOOK;
  d.comment = "";
  ASSERT_EQ(fm_workbook_set_defined_name(wb.handle, &d), 0) << fm_get_last_error_message();
  d.name = "WithNull";
  d.comment = nullptr;
  ASSERT_EQ(fm_workbook_set_defined_name(wb.handle, &d), 0) << fm_get_last_error_message();

  for (const char* name : {"WithEmpty", "WithNull"}) {
    int32_t present = 0;
    fm_defined_name_t out = FM_STRUCT_INIT(fm_defined_name_t);
    ASSERT_EQ(fm_workbook_get_defined_name(wb.handle, name, FM_DEFINED_NAME_SCOPE_WORKBOOK, &present, &out), 0);
    ASSERT_EQ(present, 1) << name;
    EXPECT_EQ(out.comment, nullptr) << name;
  }
}

// ---------------------------------------------------------------------------
// Notes
// ---------------------------------------------------------------------------

TEST(TextFieldsEmptyStrings, NoteEmptyTextWritesANoteWithNoTextAndNullRemovesIt) {
  Book wb;
  fm_comment_t note = FM_STRUCT_INIT(fm_comment_t);
  note.text = "";
  ASSERT_EQ(fm_sheet_set_comment(wb.handle, 0, 1, 1, &note), 0) << fm_get_last_error_message();
  size_t count = 99;
  ASSERT_EQ(fm_sheet_get_comment_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 1U);
  int32_t present = 0;
  fm_comment_t out = FM_STRUCT_INIT(fm_comment_t);
  ASSERT_EQ(fm_sheet_get_comment(wb.handle, 0, 1, 1, &present, &out), 0);
  ASSERT_EQ(present, 1);
  EXPECT_STREQ(out.text, "");

  // A NULL text is the same note, and the absent value for the whole comment removes it.
  note.text = nullptr;
  ASSERT_EQ(fm_sheet_set_comment(wb.handle, 0, 1, 1, &note), 0);
  ASSERT_EQ(fm_sheet_get_comment_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 1U);
  ASSERT_EQ(fm_sheet_set_comment(wb.handle, 0, 1, 1, nullptr), 0);
  ASSERT_EQ(fm_sheet_get_comment(wb.handle, 0, 1, 1, &present, &out), 0);
  EXPECT_EQ(present, 0);
  ASSERT_EQ(fm_sheet_get_comment_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST(TextFieldsEmptyStrings, NoteEmptyAuthorAndNullAreOneState) {
  Book wb;
  fm_comment_t note = FM_STRUCT_INIT(fm_comment_t);
  note.text = "a";
  note.author = "";
  ASSERT_EQ(fm_sheet_set_comment(wb.handle, 0, 1, 1, &note), 0) << fm_get_last_error_message();
  int32_t present = 0;
  fm_comment_t out = FM_STRUCT_INIT(fm_comment_t);
  ASSERT_EQ(fm_sheet_get_comment(wb.handle, 0, 1, 1, &present, &out), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(out.author, nullptr);
}

// ---------------------------------------------------------------------------
// People and threaded comments
// ---------------------------------------------------------------------------

size_t PersonCount(const fm_workbook_t* wb) {
  size_t count = 99;
  EXPECT_EQ(fm_workbook_get_person_count(wb, &count), 0);
  return count;
}

TEST(TextFieldsEmptyStrings, PersonRefusesAnEmptyIdAndProviderIdAndTakesAnEmptyUserIdAsAbsent) {
  Book wb;
  const char* id = nullptr;
  fm_person_t person = FM_STRUCT_INIT(fm_person_t);
  person.display_name = "Ann";
  person.id = "";
  EXPECT_EQ(fm_workbook_add_person(wb.handle, &person, &id), kInvalidArgument);
  person.id = nullptr;
  person.provider_id = "";
  EXPECT_EQ(fm_workbook_add_person(wb.handle, &person, &id), kInvalidArgument);
  EXPECT_EQ(PersonCount(wb.handle), 0U);

  person.provider_id = nullptr;
  person.user_id = "";
  ASSERT_EQ(fm_workbook_add_person(wb.handle, &person, &id), 0) << fm_get_last_error_message();
  fm_person_t out = FM_STRUCT_INIT(fm_person_t);
  ASSERT_EQ(fm_workbook_get_person_at(wb.handle, 0, &out), 0);
  EXPECT_EQ(out.user_id, nullptr);
  EXPECT_EQ(out.provider_id, nullptr);
  EXPECT_NE(out.id, nullptr);
}

TEST(TextFieldsEmptyStrings, ThreadedCommentRefusesEmptyIdCreatedAndMentionIdsAndNullGenerates) {
  Book wb;
  const char* stored = nullptr;
  fm_person_t person = FM_STRUCT_INIT(fm_person_t);
  person.display_name = "Ann";
  ASSERT_EQ(fm_workbook_add_person(wb.handle, &person, &stored), 0) << fm_get_last_error_message();
  const std::string person_id = stored;

  fm_threaded_comment_t c = FM_STRUCT_INIT(fm_threaded_comment_t);
  c.row = 1;
  c.col = 1;
  c.person_id = person_id.c_str();
  c.text = "@Ann hi";
  const char* out_id = nullptr;

  c.id = "";
  EXPECT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, FM_TEXT_OFFSET_UNIT_UTF16, &c, &out_id), kInvalidArgument);
  c.id = nullptr;
  c.created = "";
  EXPECT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, FM_TEXT_OFFSET_UNIT_UTF16, &c, &out_id), kInvalidArgument);
  c.created = nullptr;

  fm_mention_t mention = {person_id.c_str(), nullptr, 0, 4, 0, 4};
  c.mentions = &mention;
  c.mentions_count = 1;
  mention.person_id = "";
  EXPECT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, FM_TEXT_OFFSET_UNIT_UTF16, &c, &out_id), kInvalidArgument);
  mention.person_id = person_id.c_str();
  mention.mention_id = "";
  EXPECT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, FM_TEXT_OFFSET_UNIT_UTF16, &c, &out_id), kInvalidArgument);
  size_t count = 99;
  ASSERT_EQ(fm_sheet_get_threaded_comment_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);

  mention.mention_id = nullptr;
  ASSERT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, FM_TEXT_OFFSET_UNIT_UTF16, &c, &out_id), 0)
      << fm_get_last_error_message();
  fm_threaded_comment_t out = FM_STRUCT_INIT(fm_threaded_comment_t);
  ASSERT_EQ(fm_sheet_get_threaded_comment_at(wb.handle, 0, 0, &out), 0);
  ASSERT_EQ(out.mentions_count, 1U);
  ASSERT_NE(out.mentions[0].mention_id, nullptr);
  EXPECT_EQ(std::string(out.mentions[0].mention_id).front(), '{');
  EXPECT_NE(out.created, nullptr);
}

// ---------------------------------------------------------------------------
// Table
// ---------------------------------------------------------------------------

size_t TableCount(const fm_workbook_t* wb) {
  size_t count = 99;
  EXPECT_EQ(fm_sheet_get_table_count(wb, 0, &count), 0);
  return count;
}

TEST(TextFieldsEmptyStrings, TableRefusesAnEmptyNameAndColumnNameAndNullDerives) {
  Book wb;
  size_t index = 0;
  fm_table_spec_t spec = FM_STRUCT_INIT(fm_table_spec_t);
  spec.range = fm_cell_rect_t{0, 0, 2, 1};
  spec.name = "";
  EXPECT_EQ(fm_sheet_add_table(wb.handle, 0, &spec, &index), kInvalidArgument);

  spec.name = "Sales";
  const char* all_empty[] = {"", ""};
  spec.column_names = all_empty;
  spec.column_count = 2;
  EXPECT_EQ(fm_sheet_add_table(wb.handle, 0, &spec, &index), kInvalidArgument);
  const char* one_empty[] = {"A", ""};
  spec.column_names = one_empty;
  EXPECT_EQ(fm_sheet_add_table(wb.handle, 0, &spec, &index), kInvalidArgument);
  EXPECT_EQ(TableCount(wb.handle), 0U);

  // A NULL entry takes the name of its header cell.
  const char* derived[] = {"A", nullptr};
  spec.column_names = derived;
  ASSERT_EQ(fm_sheet_add_table(wb.handle, 0, &spec, &index), 0) << fm_get_last_error_message();
  EXPECT_EQ(TableCount(wb.handle), 1U);
}

// ---------------------------------------------------------------------------
// Protection
// ---------------------------------------------------------------------------

TEST(TextFieldsEmptyStrings, ProtectedRangeRefusesEmptyHashSaltAlgorithmAndDescriptor) {
  struct Field {
    const char* name;
    const char* fm_protected_range_t::*member;
    const char* fm_protected_range_patch_t::*patch_member;
    fm_patch_op_t fm_protected_range_patch_t::*patch_op;
  };
  const Field fields[] = {
      {"legacy_password_hash", &fm_protected_range_t::legacy_password_hash,
       &fm_protected_range_patch_t::legacy_password_hash, &fm_protected_range_patch_t::legacy_password_hash_op},
      {"algorithm_name", &fm_protected_range_t::algorithm_name, &fm_protected_range_patch_t::algorithm_name,
       &fm_protected_range_patch_t::algorithm_name_op},
      {"hash_value", &fm_protected_range_t::hash_value, &fm_protected_range_patch_t::hash_value,
       &fm_protected_range_patch_t::hash_value_op},
      {"salt_value", &fm_protected_range_t::salt_value, &fm_protected_range_patch_t::salt_value,
       &fm_protected_range_patch_t::salt_value_op},
      {"security_descriptor", &fm_protected_range_t::security_descriptor,
       &fm_protected_range_patch_t::security_descriptor, &fm_protected_range_patch_t::security_descriptor_op},
  };
  const fm_cell_rect_t rect = {0, 0, 1, 1};
  for (const Field& f : fields) {
    SCOPED_TRACE(f.name);
    Book wb;
    fm_protected_range_t range = FM_STRUCT_INIT(fm_protected_range_t);
    range.name = "R";
    range.ranges = &rect;
    range.range_count = 1;
    range.*f.member = "";
    EXPECT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &range, nullptr), kInvalidArgument);
    size_t count = 99;
    ASSERT_EQ(fm_sheet_get_protected_range_count(wb.handle, 0, &count), 0);
    EXPECT_EQ(count, 0U);

    range.*f.member = "x";
    ASSERT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &range, nullptr), 0) << fm_get_last_error_message();
    fm_protected_range_patch_t patch = FM_STRUCT_INIT(fm_protected_range_patch_t);
    patch.*f.patch_op = FM_PATCH_OP_SET;
    patch.*f.patch_member = "";
    EXPECT_EQ(fm_sheet_update_protected_range(wb.handle, 0, "R", &patch), kInvalidArgument);
    fm_protected_range_t out = FM_STRUCT_INIT(fm_protected_range_t);
    ASSERT_EQ(fm_sheet_get_protected_range_at(wb.handle, 0, 0, &out), 0);
    EXPECT_STREQ(out.*f.member, "x");

    patch.*f.patch_op = FM_PATCH_OP_CLEAR;
    patch.*f.patch_member = nullptr;
    ASSERT_EQ(fm_sheet_update_protected_range(wb.handle, 0, "R", &patch), 0) << fm_get_last_error_message();
    ASSERT_EQ(fm_sheet_get_protected_range_at(wb.handle, 0, 0, &out), 0);
    EXPECT_EQ(out.*f.member, nullptr);
  }
}

TEST(TextFieldsEmptyStrings, SheetProtectionRefusesEmptyStringsOnSetAndUpdate) {
  struct Field {
    const char* name;
    const char* fm_sheet_protection_t::*member;
    const char* fm_sheet_protection_patch_t::*patch_member;
    fm_patch_op_t fm_sheet_protection_patch_t::*patch_op;
  };
  const Field fields[] = {
      {"algorithm_name", &fm_sheet_protection_t::algorithm_name, &fm_sheet_protection_patch_t::algorithm_name,
       &fm_sheet_protection_patch_t::algorithm_name_op},
      {"hash_value", &fm_sheet_protection_t::hash_value, &fm_sheet_protection_patch_t::hash_value,
       &fm_sheet_protection_patch_t::hash_value_op},
      {"salt_value", &fm_sheet_protection_t::salt_value, &fm_sheet_protection_patch_t::salt_value,
       &fm_sheet_protection_patch_t::salt_value_op},
      {"legacy_password_hash", &fm_sheet_protection_t::legacy_password_hash,
       &fm_sheet_protection_patch_t::legacy_password_hash, &fm_sheet_protection_patch_t::legacy_password_hash_op},
  };
  for (const Field& f : fields) {
    SCOPED_TRACE(f.name);
    Book wb;
    fm_sheet_protection_t in = FM_STRUCT_INIT(fm_sheet_protection_t);
    in.*f.member = "";
    EXPECT_EQ(fm_sheet_set_sheet_protection(wb.handle, 0, &in), kInvalidArgument);
    int32_t present = 1;
    fm_sheet_protection_t out = FM_STRUCT_INIT(fm_sheet_protection_t);
    ASSERT_EQ(fm_sheet_get_sheet_protection(wb.handle, 0, &present, &out), 0);
    EXPECT_EQ(present, 0);

    in.*f.member = "x";
    ASSERT_EQ(fm_sheet_set_sheet_protection(wb.handle, 0, &in), 0) << fm_get_last_error_message();
    fm_sheet_protection_patch_t patch = FM_STRUCT_INIT(fm_sheet_protection_patch_t);
    patch.*f.patch_op = FM_PATCH_OP_SET;
    patch.*f.patch_member = "";
    EXPECT_EQ(fm_sheet_update_sheet_protection(wb.handle, 0, &patch), kInvalidArgument);
    ASSERT_EQ(fm_sheet_get_sheet_protection(wb.handle, 0, &present, &out), 0);
    EXPECT_STREQ(out.*f.member, "x");

    patch.*f.patch_op = FM_PATCH_OP_CLEAR;
    patch.*f.patch_member = nullptr;
    ASSERT_EQ(fm_sheet_update_sheet_protection(wb.handle, 0, &patch), 0) << fm_get_last_error_message();
    ASSERT_EQ(fm_sheet_get_sheet_protection(wb.handle, 0, &present, &out), 0);
    EXPECT_EQ(out.*f.member, nullptr);
  }
}

TEST(TextFieldsEmptyStrings, WorkbookProtectionRefusesEmptyStringsAndNullIsAbsent) {
  const char* fm_workbook_protection_t::*const members[] = {
      &fm_workbook_protection_t::workbook_legacy_password_hash,
      &fm_workbook_protection_t::workbook_algorithm_name,
      &fm_workbook_protection_t::workbook_hash_value,
      &fm_workbook_protection_t::workbook_salt_value,
      &fm_workbook_protection_t::revisions_legacy_password_hash,
      &fm_workbook_protection_t::revisions_algorithm_name,
      &fm_workbook_protection_t::revisions_hash_value,
      &fm_workbook_protection_t::revisions_salt_value,
  };
  for (std::size_t i = 0; i < sizeof(members) / sizeof(members[0]); ++i) {
    SCOPED_TRACE(i);
    Book wb;
    fm_workbook_protection_t p = FM_STRUCT_INIT(fm_workbook_protection_t);
    p.lock_structure = 1;
    p.*members[i] = "";
    EXPECT_EQ(fm_workbook_set_workbook_protection(wb.handle, &p), kInvalidArgument);
    int32_t present = 1;
    fm_workbook_protection_t out = FM_STRUCT_INIT(fm_workbook_protection_t);
    ASSERT_EQ(fm_workbook_get_workbook_protection(wb.handle, &present, &out), 0);
    EXPECT_EQ(present, 0);

    p.*members[i] = nullptr;
    ASSERT_EQ(fm_workbook_set_workbook_protection(wb.handle, &p), 0) << fm_get_last_error_message();
    ASSERT_EQ(fm_workbook_get_workbook_protection(wb.handle, &present, &out), 0);
    EXPECT_EQ(present, 1);
    EXPECT_EQ(out.*members[i], nullptr);
  }
}

// ---------------------------------------------------------------------------
// Differential number format
// ---------------------------------------------------------------------------

TEST(TextFieldsEmptyStrings, DxfEmptyNumberFormatIsTheAbsentValue) {
  Book wb;
  fm_dxf_t dxf{};
  dxf.has_font = 1;
  dxf.font.has_bold = 1;
  dxf.font.bold = 1;
  dxf.number_format = "";
  uint32_t index = 99;
  ASSERT_EQ(fm_workbook_add_dxf(wb.handle, &dxf, &index), 0) << fm_get_last_error_message();
  fm_dxf_t out{};
  ASSERT_EQ(fm_workbook_get_dxf_at(wb.handle, index, &out), 0);
  EXPECT_EQ(out.number_format, nullptr);
}

}  // namespace
