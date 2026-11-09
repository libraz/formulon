// Stable C ABI number-format tests.

#include "formulon_c_styles_test_helpers.h"

TEST(FormulonCApiStyles, BuiltinNumFmtResolves) {
  int32_t present = 0;
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const char* s = nullptr;
  EXPECT_EQ(fm_workbook_get_number_format(wb.handle, 0, &present, &s), 0);
  ASSERT_NE(s, nullptr);
  EXPECT_STREQ(s, "General");
  EXPECT_EQ(fm_workbook_get_number_format(wb.handle, 14, &present, &s), 0);
  ASSERT_NE(s, nullptr);
  EXPECT_STREQ(s, "m/d/yy");
}

TEST(FormulonCApiStyles, AddNumberFormatUsesProfileVocabulary) {
  int32_t present = 0;
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  uint32_t id = 0xFFFFU;
  ASSERT_EQ(fm_workbook_add_number_format(wb.handle, "m/d/yy", &id), 0);
  EXPECT_EQ(id, 14U);
  const char* s = nullptr;
  ASSERT_EQ(fm_workbook_get_number_format(wb.handle, id, &present, &s), 0);
  EXPECT_STREQ(s, "m/d/yy");

  // The ECMA-376 spelling is not the profile's code, so it is a custom format.
  ASSERT_EQ(fm_workbook_add_number_format(wb.handle, "mm-dd-yy", &id), 0);
  EXPECT_GE(id, 164U);
  ASSERT_EQ(fm_workbook_get_number_format(wb.handle, id, &present, &s), 0);
  EXPECT_STREQ(s, "mm-dd-yy");

  ASSERT_EQ(fm_workbook_set_profile_id(wb.handle, "mac-365-ja_JP"), 0);
  ASSERT_EQ(fm_workbook_add_number_format(wb.handle, "yyyy/m/d", &id), 0);
  EXPECT_EQ(id, 14U);
  ASSERT_EQ(fm_workbook_get_number_format(wb.handle, 14U, &present, &s), 0);
  EXPECT_STREQ(s, "yyyy/m/d");
}

TEST(FormulonCApiStyles, LocaleOnlyBuiltinNumFmtFollowsProfile) {
  int32_t present = 0;
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const char* s = nullptr;
  // Under en-US the East Asian date ids render as the short date.
  ASSERT_EQ(fm_workbook_get_number_format(wb.handle, 31, &present, &s), 0);
  EXPECT_STREQ(s, "m/d/yy");

  ASSERT_EQ(fm_workbook_set_profile_id(wb.handle, "mac-365-ja_JP"), 0);
  ASSERT_EQ(fm_workbook_get_number_format(wb.handle, 31, &present, &s), 0);
  EXPECT_STREQ(s, "yyyy\"年\"m\"月\"d\"日\"");
  ASSERT_EQ(fm_workbook_get_number_format(wb.handle, 57, &present, &s), 0);
  EXPECT_STREQ(s, "[$-411]ge\\.m\\.d");
  ASSERT_EQ(fm_workbook_get_number_format(wb.handle, 14, &present, &s), 0);
  EXPECT_STREQ(s, "yyyy/m/d");

  ASSERT_EQ(fm_workbook_set_profile_id(wb.handle, "mac-365-zh_CN"), 0);
  ASSERT_EQ(fm_workbook_get_number_format(wb.handle, 57, &present, &s), 0);
  EXPECT_STREQ(s, "yyyy\"年\"m\"月\"");

  ASSERT_EQ(fm_workbook_set_profile_id(wb.handle, "mac-365-th_TH"), 0);
  ASSERT_EQ(fm_workbook_get_number_format(wb.handle, 59, &present, &s), 0);
  EXPECT_STREQ(s, "t0");
}

TEST(FormulonCApiStyles, CustomNumFmtOverridingBuiltinIdWins) {
  int32_t present = 0;
  // A file may define a custom <numFmt> whose numFmtId collides with a
  // built-in slot; Excel honours the file's definition. Parse a production-
  // shaped styles.xml and install the resulting table rather than manually
  // constructing the record.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadNumFmtOverrideStyles(wb.handle));

  const char* s = nullptr;
  ASSERT_EQ(fm_workbook_get_number_format(wb.handle, 14, &present, &s), 0);
  ASSERT_NE(s, nullptr);
  EXPECT_STREQ(s, "yyyy");
}

TEST(FormulonCApiStyles, AddNumFmtUsesEffectiveBuiltinMapping) {
  int32_t present = 0;
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadNumFmtOverrideStyles(wb.handle));

  const char* resolved = nullptr;
  ASSERT_EQ(fm_workbook_get_number_format(wb.handle, 14U, &present, &resolved), 0);
  ASSERT_NE(resolved, nullptr);
  EXPECT_STREQ(resolved, "yyyy");

  uint32_t override_id = 0xFFFFU;
  ASSERT_EQ(fm_workbook_add_number_format(wb.handle, "yyyy", &override_id), 0);
  EXPECT_EQ(override_id, 14U);

  uint32_t builtin_code_id = 0xFFFFU;
  ASSERT_EQ(fm_workbook_add_number_format(wb.handle, "mm-dd-yy", &builtin_code_id), 0);
  EXPECT_GE(builtin_code_id, 164U);
  ASSERT_EQ(fm_workbook_get_number_format(wb.handle, builtin_code_id, &present, &resolved), 0);
  ASSERT_NE(resolved, nullptr);
  EXPECT_STREQ(resolved, "mm-dd-yy");

  uint32_t builtin_code_again = 0xFFFFU;
  ASSERT_EQ(fm_workbook_add_number_format(wb.handle, "mm-dd-yy", &builtin_code_again), 0);
  EXPECT_EQ(builtin_code_again, builtin_code_id);
}

TEST(FormulonCApiStyles, AddBatchNumFmtUsesEffectiveBuiltinMapping) {
  int32_t present = 0;
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadNumFmtOverrideStyles(wb.handle));

  const char* const codes[] = {"General", "yyyy", "mm-dd-yy", "mm-dd-yy"};
  uint32_t ids[] = {0xFFFFU, 0xFFFFU, 0xFFFFU, 0xFFFFU};
  fm_styles_batch_spec_t batch = FM_STRUCT_INIT(fm_styles_batch_spec_t);
  batch.num_fmt_codes = codes;
  batch.num_fmt_count = 4U;
  batch.num_fmt_ids = ids;

  ASSERT_EQ(fm_workbook_add_batch(wb.handle, &batch), 0);
  EXPECT_EQ(ids[0], 0U);
  EXPECT_EQ(ids[1], 14U);
  EXPECT_GE(ids[2], 164U);
  EXPECT_EQ(ids[3], ids[2]);
  for (size_t i = 0; i < 4U; ++i) {
    const char* resolved = nullptr;
    ASSERT_EQ(fm_workbook_get_number_format(wb.handle, ids[i], &present, &resolved), 0);
    ASSERT_NE(resolved, nullptr);
    EXPECT_STREQ(resolved, codes[i]);
  }
}

TEST(FormulonCApiStyles, InvalidCustomNumFmtStringIndexDoesNotShadowBuiltin) {
  int32_t present = 0;
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadNumFmtOverrideStyles(wb.handle));

  auto& styles = wb.handle->workbook().mutable_styles();
  ASSERT_EQ(styles.num_fmts.size(), 1U);
  styles.num_fmts[0].format_string_index = static_cast<std::uint32_t>(styles.num_fmt_strings.size());

  const char* resolved = nullptr;
  ASSERT_EQ(fm_workbook_get_number_format(wb.handle, 14U, &present, &resolved), 0);
  ASSERT_NE(resolved, nullptr);
  EXPECT_STREQ(resolved, "m/d/yy");

  uint32_t id = 0xFFFFU;
  ASSERT_EQ(fm_workbook_add_number_format(wb.handle, "m/d/yy", &id), 0);
  EXPECT_EQ(id, 14U);
}

TEST(FormulonCApiStyles, DuplicateNumFmtFirstValidRecordWins) {
  int32_t present = 0;
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadNumFmtDuplicateStyles(wb.handle));

  const char* resolved = nullptr;
  ASSERT_EQ(fm_workbook_get_number_format(wb.handle, 14U, &present, &resolved), 0);
  ASSERT_NE(resolved, nullptr);
  EXPECT_STREQ(resolved, "yyyy");

  uint32_t first_id = 0xFFFFU;
  ASSERT_EQ(fm_workbook_add_number_format(wb.handle, "yyyy", &first_id), 0);
  EXPECT_EQ(first_id, 14U);

  uint32_t shadowed_id = 0xFFFFU;
  ASSERT_EQ(fm_workbook_add_number_format(wb.handle, "dd/mm/yyyy", &shadowed_id), 0);
  EXPECT_GE(shadowed_id, 164U);
  ASSERT_EQ(fm_workbook_get_number_format(wb.handle, shadowed_id, &present, &resolved), 0);
  ASSERT_NE(resolved, nullptr);
  EXPECT_STREQ(resolved, "dd/mm/yyyy");
}

TEST(FormulonCApiStyles, DuplicateNumFmtInvalidFirstThenValidRecordWins) {
  int32_t present = 0;
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadNumFmtDuplicateStyles(wb.handle));

  auto& styles = wb.handle->workbook().mutable_styles();
  ASSERT_EQ(styles.num_fmts.size(), 2U);
  styles.num_fmts[0].format_string_index = static_cast<std::uint32_t>(styles.num_fmt_strings.size());

  const char* resolved = nullptr;
  ASSERT_EQ(fm_workbook_get_number_format(wb.handle, 14U, &present, &resolved), 0);
  ASSERT_NE(resolved, nullptr);
  EXPECT_STREQ(resolved, "dd/mm/yyyy");

  uint32_t valid_id = 0xFFFFU;
  ASSERT_EQ(fm_workbook_add_number_format(wb.handle, "dd/mm/yyyy", &valid_id), 0);
  EXPECT_EQ(valid_id, 14U);

  uint32_t invalid_id = 0xFFFFU;
  ASSERT_EQ(fm_workbook_add_number_format(wb.handle, "yyyy", &invalid_id), 0);
  EXPECT_GE(invalid_id, 164U);
  ASSERT_EQ(fm_workbook_get_number_format(wb.handle, invalid_id, &present, &resolved), 0);
  ASSERT_NE(resolved, nullptr);
  EXPECT_STREQ(resolved, "yyyy");
}

TEST(FormulonCApiStyles, DuplicateNumFmtAllInvalidRecordsFallBackToBuiltin) {
  int32_t present = 0;
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadNumFmtDuplicateStyles(wb.handle));

  auto& styles = wb.handle->workbook().mutable_styles();
  ASSERT_EQ(styles.num_fmts.size(), 2U);
  const std::uint32_t invalid_index = static_cast<std::uint32_t>(styles.num_fmt_strings.size());
  for (auto& record : styles.num_fmts) {
    record.format_string_index = invalid_index;
  }

  const char* resolved = nullptr;
  ASSERT_EQ(fm_workbook_get_number_format(wb.handle, 14U, &present, &resolved), 0);
  ASSERT_NE(resolved, nullptr);
  EXPECT_STREQ(resolved, "m/d/yy");

  uint32_t builtin_id = 0xFFFFU;
  ASSERT_EQ(fm_workbook_add_number_format(wb.handle, "m/d/yy", &builtin_id), 0);
  EXPECT_EQ(builtin_id, 14U);
}

TEST(FormulonCApiStyles, UnknownNumFmtIdIsAbsent) {
  int32_t present = 0;
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const char* s = nullptr;
  // A built-in slot no locale assigns a code is absent rather than an empty
  // string.
  ASSERT_EQ(fm_workbook_get_number_format(wb.handle, 82, &present, &s), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(s, nullptr);
  EXPECT_EQ(fm_workbook_get_number_format(wb.handle, 82, nullptr, &s),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  // Id 5 is reserved in ECMA-376 but rendered as the locale currency.
  ASSERT_EQ(fm_workbook_get_number_format(wb.handle, 5, &present, &s), 0);
  EXPECT_EQ(present, 1);
  EXPECT_STREQ(s, "\"$\"#,##0_);\\(\"$\"#,##0\\)");
}

TEST(FormulonCApiStyles, AddNumFmtBuiltinReturnsBuiltinId) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  uint32_t id = 0xFFFFU;
  ASSERT_EQ(fm_workbook_add_number_format(wb.handle, "General", &id), 0);
  EXPECT_EQ(id, 0U);
  // Adding a built-in must not create a custom entry.
  std::size_t font_count = 7;  // unrelated, just ensure other tables untouched
  EXPECT_EQ(fm_workbook_get_font_count(wb.handle, &font_count), 0);
  EXPECT_EQ(font_count, 1U);  // the seeded default only
}

TEST(FormulonCApiStyles, AddNumFmtCustomReturnsCustomId) {
  int32_t present = 0;
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  uint32_t id = 0;
  ASSERT_EQ(fm_workbook_add_number_format(wb.handle, "\"USD\" #,##0", &id), 0);
  EXPECT_GE(id, 164U);
  // Resolves through the read-side getter.
  const char* s = nullptr;
  ASSERT_EQ(fm_workbook_get_number_format(wb.handle, id, &present, &s), 0);
  ASSERT_NE(s, nullptr);
  EXPECT_STREQ(s, "\"USD\" #,##0");
}

TEST(FormulonCApiStyles, AddNumFmtCustomDedup) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  uint32_t a = 0;
  uint32_t b = 0;
  ASSERT_EQ(fm_workbook_add_number_format(wb.handle, "\"USD\" #,##0", &a), 0);
  ASSERT_EQ(fm_workbook_add_number_format(wb.handle, "\"USD\" #,##0", &b), 0);
  EXPECT_EQ(a, b);
}

TEST(FormulonCApiStyles, AddNumFmtExhaustionReturnsPreconditionWithoutMutation) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  auto& styles = wb.handle->workbook().mutable_styles();
  formulon::NumFmtRecord max_record;
  max_record.id = 65535U;
  max_record.format_string_index = static_cast<std::uint32_t>(styles.num_fmt_strings.size());
  styles.num_fmt_strings.emplace_back("existing");
  styles.num_fmts.push_back(max_record);
  const std::size_t strings_before = styles.num_fmt_strings.size();
  const std::size_t records_before = styles.num_fmts.size();
  uint32_t out = 0xBEEFU;
  EXPECT_EQ(fm_workbook_add_number_format(wb.handle, "new-format", &out),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kPreconditionFailed));
  EXPECT_EQ(out, 0xBEEFU);
  EXPECT_EQ(styles.num_fmt_strings.size(), strings_before);
  EXPECT_EQ(styles.num_fmts.size(), records_before);
}

TEST(FormulonCApiStyles, AddBatchNumFmtExhaustionIsTransactional) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  auto& styles = wb.handle->workbook().mutable_styles();
  formulon::NumFmtRecord max_record;
  max_record.id = 65535U;
  max_record.format_string_index = static_cast<std::uint32_t>(styles.num_fmt_strings.size());
  styles.num_fmt_strings.emplace_back("existing");
  styles.num_fmts.push_back(max_record);
  const std::string before_xml = formulon::io::write_styles(styles);
  const char* codes[] = {"first", "second"};
  uint32_t ids[] = {0xAAAAU, 0xBBBBU};
  fm_styles_batch_spec_t batch = FM_STRUCT_INIT(fm_styles_batch_spec_t);
  batch.num_fmt_codes = codes;
  batch.num_fmt_count = 2U;
  batch.num_fmt_ids = ids;
  EXPECT_EQ(fm_workbook_add_batch(wb.handle, &batch),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kPreconditionFailed));
  EXPECT_EQ(ids[0], 0xAAAAU);
  EXPECT_EQ(ids[1], 0xBBBBU);
  EXPECT_EQ(formulon::io::write_styles(styles), before_xml);
}

TEST(FormulonCApiNumberFormatList, ListsDeclaredFormatsInIdOrder) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t count = 99;
  ASSERT_EQ(fm_workbook_get_number_format_count(wb.handle, &count), 0);
  const std::size_t base = count;
  std::uint32_t first = 0;
  std::uint32_t second = 0;
  ASSERT_EQ(fm_workbook_add_number_format(wb.handle, "0.000\"kg\"", &first), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_add_number_format(wb.handle, "#,##0.0000", &second), 0) << fm_get_last_error_message();
  ASSERT_LT(first, second);
  // A built-in code is not declared, so it does not join the list.
  std::uint32_t builtin = 0;
  ASSERT_EQ(fm_workbook_add_number_format(wb.handle, "0.00", &builtin), 0);
  EXPECT_LT(builtin, 164U);
  ASSERT_EQ(fm_workbook_get_number_format_count(wb.handle, &count), 0);
  ASSERT_EQ(count, base + 2U);

  std::uint32_t previous = 0;
  bool saw_first = false;
  bool saw_second = false;
  for (std::size_t i = 0; i < count; ++i) {
    fm_number_format_info_t info = FM_STRUCT_INIT(fm_number_format_info_t);
    ASSERT_EQ(fm_workbook_get_number_format_at(wb.handle, i, &info), 0) << fm_get_last_error_message();
    ASSERT_NE(info.number_format, nullptr);
    if (i > 0) {
      EXPECT_GT(info.num_fmt_id, previous);
    }
    previous = info.num_fmt_id;
    if (info.num_fmt_id == first) {
      EXPECT_STREQ(info.number_format, "0.000\"kg\"");
      saw_first = true;
    }
    if (info.num_fmt_id == second) {
      EXPECT_STREQ(info.number_format, "#,##0.0000");
      saw_second = true;
    }
  }
  EXPECT_TRUE(saw_first);
  EXPECT_TRUE(saw_second);
  fm_number_format_info_t past = FM_STRUCT_INIT(fm_number_format_info_t);
  EXPECT_EQ(fm_workbook_get_number_format_at(wb.handle, count, &past), FM_STATUS_INDEX_OUT_OF_RANGE);
}
