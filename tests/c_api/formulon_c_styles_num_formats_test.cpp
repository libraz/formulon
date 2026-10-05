// Stable C ABI number-format tests.

#include "formulon_c_styles_test_helpers.h"

TEST(FormulonCApiStyles, BuiltinNumFmtResolves) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const char* s = nullptr;
  EXPECT_EQ(fm_styles_get_num_fmt_string(wb.handle, 0, &s), 0);
  ASSERT_NE(s, nullptr);
  EXPECT_STREQ(s, "General");
  EXPECT_EQ(fm_styles_get_num_fmt_string(wb.handle, 14, &s), 0);
  ASSERT_NE(s, nullptr);
  EXPECT_STREQ(s, "mm-dd-yy");
}

TEST(FormulonCApiStyles, CustomNumFmtOverridingBuiltinIdWins) {
  // A file may define a custom <numFmt> whose numFmtId collides with a
  // built-in slot; Excel honours the file's definition. Parse a production-
  // shaped styles.xml and install the resulting table rather than manually
  // constructing the record.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadNumFmtOverrideStyles(wb.handle));

  const char* s = nullptr;
  ASSERT_EQ(fm_styles_get_num_fmt_string(wb.handle, 14, &s), 0);
  ASSERT_NE(s, nullptr);
  EXPECT_STREQ(s, "yyyy");
}

TEST(FormulonCApiStyles, AddNumFmtUsesEffectiveBuiltinMapping) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadNumFmtOverrideStyles(wb.handle));

  const char* resolved = nullptr;
  ASSERT_EQ(fm_styles_get_num_fmt_string(wb.handle, 14U, &resolved), 0);
  ASSERT_NE(resolved, nullptr);
  EXPECT_STREQ(resolved, "yyyy");

  uint16_t override_id = 0xFFFFU;
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "yyyy", &override_id), 0);
  EXPECT_EQ(override_id, 14U);

  uint16_t builtin_code_id = 0xFFFFU;
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "mm-dd-yy", &builtin_code_id), 0);
  EXPECT_GE(builtin_code_id, 164U);
  ASSERT_EQ(fm_styles_get_num_fmt_string(wb.handle, builtin_code_id, &resolved), 0);
  ASSERT_NE(resolved, nullptr);
  EXPECT_STREQ(resolved, "mm-dd-yy");

  uint16_t builtin_code_again = 0xFFFFU;
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "mm-dd-yy", &builtin_code_again), 0);
  EXPECT_EQ(builtin_code_again, builtin_code_id);
}

TEST(FormulonCApiStyles, AddBatchNumFmtUsesEffectiveBuiltinMapping) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadNumFmtOverrideStyles(wb.handle));

  const char* const codes[] = {"General", "yyyy", "mm-dd-yy", "mm-dd-yy"};
  uint16_t ids[] = {0xFFFFU, 0xFFFFU, 0xFFFFU, 0xFFFFU};
  fm_styles_batch batch{};
  batch.num_fmt_codes = codes;
  batch.num_fmt_count = 4U;
  batch.num_fmt_ids = ids;

  ASSERT_EQ(fm_styles_add_batch(wb.handle, &batch), 0);
  EXPECT_EQ(ids[0], 0U);
  EXPECT_EQ(ids[1], 14U);
  EXPECT_GE(ids[2], 164U);
  EXPECT_EQ(ids[3], ids[2]);
  for (size_t i = 0; i < 4U; ++i) {
    const char* resolved = nullptr;
    ASSERT_EQ(fm_styles_get_num_fmt_string(wb.handle, ids[i], &resolved), 0);
    ASSERT_NE(resolved, nullptr);
    EXPECT_STREQ(resolved, codes[i]);
  }
}

TEST(FormulonCApiStyles, InvalidCustomNumFmtStringIndexDoesNotShadowBuiltin) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadNumFmtOverrideStyles(wb.handle));

  auto& styles = wb.handle->workbook().mutable_styles();
  ASSERT_EQ(styles.num_fmts.size(), 1U);
  styles.num_fmts[0].format_string_index = static_cast<std::uint32_t>(styles.num_fmt_strings.size());

  const char* resolved = nullptr;
  ASSERT_EQ(fm_styles_get_num_fmt_string(wb.handle, 14U, &resolved), 0);
  ASSERT_NE(resolved, nullptr);
  EXPECT_STREQ(resolved, "mm-dd-yy");

  uint16_t id = 0xFFFFU;
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "mm-dd-yy", &id), 0);
  EXPECT_EQ(id, 14U);
}

TEST(FormulonCApiStyles, DuplicateNumFmtFirstValidRecordWins) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadNumFmtDuplicateStyles(wb.handle));

  const char* resolved = nullptr;
  ASSERT_EQ(fm_styles_get_num_fmt_string(wb.handle, 14U, &resolved), 0);
  ASSERT_NE(resolved, nullptr);
  EXPECT_STREQ(resolved, "yyyy");

  uint16_t first_id = 0xFFFFU;
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "yyyy", &first_id), 0);
  EXPECT_EQ(first_id, 14U);

  uint16_t shadowed_id = 0xFFFFU;
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "dd/mm/yyyy", &shadowed_id), 0);
  EXPECT_GE(shadowed_id, 164U);
  ASSERT_EQ(fm_styles_get_num_fmt_string(wb.handle, shadowed_id, &resolved), 0);
  ASSERT_NE(resolved, nullptr);
  EXPECT_STREQ(resolved, "dd/mm/yyyy");
}

TEST(FormulonCApiStyles, DuplicateNumFmtInvalidFirstThenValidRecordWins) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadNumFmtDuplicateStyles(wb.handle));

  auto& styles = wb.handle->workbook().mutable_styles();
  ASSERT_EQ(styles.num_fmts.size(), 2U);
  styles.num_fmts[0].format_string_index = static_cast<std::uint32_t>(styles.num_fmt_strings.size());

  const char* resolved = nullptr;
  ASSERT_EQ(fm_styles_get_num_fmt_string(wb.handle, 14U, &resolved), 0);
  ASSERT_NE(resolved, nullptr);
  EXPECT_STREQ(resolved, "dd/mm/yyyy");

  uint16_t valid_id = 0xFFFFU;
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "dd/mm/yyyy", &valid_id), 0);
  EXPECT_EQ(valid_id, 14U);

  uint16_t invalid_id = 0xFFFFU;
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "yyyy", &invalid_id), 0);
  EXPECT_GE(invalid_id, 164U);
  ASSERT_EQ(fm_styles_get_num_fmt_string(wb.handle, invalid_id, &resolved), 0);
  ASSERT_NE(resolved, nullptr);
  EXPECT_STREQ(resolved, "yyyy");
}

TEST(FormulonCApiStyles, DuplicateNumFmtAllInvalidRecordsFallBackToBuiltin) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadNumFmtDuplicateStyles(wb.handle));

  auto& styles = wb.handle->workbook().mutable_styles();
  ASSERT_EQ(styles.num_fmts.size(), 2U);
  const std::uint32_t invalid_index = static_cast<std::uint32_t>(styles.num_fmt_strings.size());
  for (auto& record : styles.num_fmts) {
    record.format_string_index = invalid_index;
  }

  const char* resolved = nullptr;
  ASSERT_EQ(fm_styles_get_num_fmt_string(wb.handle, 14U, &resolved), 0);
  ASSERT_NE(resolved, nullptr);
  EXPECT_STREQ(resolved, "mm-dd-yy");

  uint16_t builtin_id = 0xFFFFU;
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "mm-dd-yy", &builtin_id), 0);
  EXPECT_EQ(builtin_id, 14U);
}

TEST(FormulonCApiStyles, UnknownNumFmtIdRejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const char* s = nullptr;
  // Reserved built-in slot (id 5) without a custom override surfaces
  // an error rather than an empty string.
  EXPECT_NE(fm_styles_get_num_fmt_string(wb.handle, 5, &s), 0);
}

TEST(FormulonCApiStyles, AddNumFmtBuiltinReturnsBuiltinId) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint16_t id = 0xFFFFU;
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "General", &id), 0);
  EXPECT_EQ(id, 0U);
  // Adding a built-in must not create a custom entry.
  uint32_t font_count = 7;  // unrelated, just ensure other tables untouched
  EXPECT_EQ(fm_styles_get_font_count(wb.handle, &font_count), 0);
  EXPECT_EQ(font_count, 1U);  // the seeded default only
}

TEST(FormulonCApiStyles, AddNumFmtCustomReturnsCustomId) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint16_t id = 0;
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "\"USD\" #,##0", &id), 0);
  EXPECT_GE(id, 164U);
  // Resolves through the read-side getter.
  const char* s = nullptr;
  ASSERT_EQ(fm_styles_get_num_fmt_string(wb.handle, id, &s), 0);
  ASSERT_NE(s, nullptr);
  EXPECT_STREQ(s, "\"USD\" #,##0");
}

TEST(FormulonCApiStyles, AddNumFmtCustomDedup) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint16_t a = 0;
  uint16_t b = 0;
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "\"USD\" #,##0", &a), 0);
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "\"USD\" #,##0", &b), 0);
  EXPECT_EQ(a, b);
}

TEST(FormulonCApiStyles, AddNumFmtExhaustionReturnsPreconditionWithoutMutation) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  auto& styles = wb.handle->workbook().mutable_styles();
  formulon::NumFmtRecord max_record;
  max_record.id = 65535U;
  max_record.format_string_index = static_cast<std::uint32_t>(styles.num_fmt_strings.size());
  styles.num_fmt_strings.emplace_back("existing");
  styles.num_fmts.push_back(max_record);
  const std::size_t strings_before = styles.num_fmt_strings.size();
  const std::size_t records_before = styles.num_fmts.size();
  uint16_t out = 0xBEEFU;
  EXPECT_EQ(fm_styles_add_num_fmt(wb.handle, "new-format", &out),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kPreconditionFailed));
  EXPECT_EQ(out, 0xBEEFU);
  EXPECT_EQ(styles.num_fmt_strings.size(), strings_before);
  EXPECT_EQ(styles.num_fmts.size(), records_before);
}

TEST(FormulonCApiStyles, AddBatchNumFmtExhaustionIsTransactional) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  auto& styles = wb.handle->workbook().mutable_styles();
  formulon::NumFmtRecord max_record;
  max_record.id = 65535U;
  max_record.format_string_index = static_cast<std::uint32_t>(styles.num_fmt_strings.size());
  styles.num_fmt_strings.emplace_back("existing");
  styles.num_fmts.push_back(max_record);
  const std::string before_xml = formulon::io::write_styles(styles);
  const char* codes[] = {"first", "second"};
  uint16_t ids[] = {0xAAAAU, 0xBBBBU};
  const fm_styles_batch batch{nullptr, 0U,      nullptr, nullptr, 0U,    nullptr, nullptr, 0U,
                              nullptr, nullptr, 0U,      nullptr, codes, 2U,      ids};
  EXPECT_EQ(fm_styles_add_batch(wb.handle, &batch),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kPreconditionFailed));
  EXPECT_EQ(ids[0], 0xAAAAU);
  EXPECT_EQ(ids[1], 0xBBBBU);
  EXPECT_EQ(formulon::io::write_styles(styles), before_xml);
}
