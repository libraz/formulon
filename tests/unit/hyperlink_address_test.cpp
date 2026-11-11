#include "hyperlink_address.h"

#include <string>

#include "gtest/gtest.h"

namespace formulon {
namespace {

struct AddressRow {
  const char* input;
  const char* target;
  const char* fragment;
};

TEST(HyperlinkAddress, NormalisesEveryMeasuredAddress) {
  const AddressRow rows[] = {
      {"https://e.x", "https://e.x/", ""},
      {"https://e.x/", "https://e.x/", ""},
      {"HTTPS://E.X/Path", "https://e.x/Path", ""},
      {"https://e.x/a b", "https://e.x/a b", ""},
      {"https://e.x/%7e", "https://e.x/~", ""},
      {"https://e.x?q=1", "https://e.x/?q=1", ""},
      {"https://e.x/b#frag", "https://e.x/b", "frag"},
      {"ftp://e.x/f", "ftp://e.x/f", ""},
      {"mailto:a@b.c", "mailto:a@b.c", ""},
      {"MAILTO:A@B.C", "mailto:A@B.C", ""},
      {"docs/a.xlsx", "docs/a.xlsx", ""},
      {"/Users/libraz/Documents/f.xlsx", "file:///Users/libraz/Documents/f.xlsx", ""},
      {"\\\\server\\share\\f.xlsx", "file:////server/share/f.xlsx", ""},
      {"https://e.x/already/normal", "https://e.x/already/normal", ""},
      {"FTP://E.X/F", "ftp://e.x/F", ""},
      {"ftp://e.x", "ftp://e.x/", ""},
      {"ftp://e.x?q=1", "ftp://e.x/?q=1", ""},
      {"ftp://e.x/%7e", "ftp://e.x/~", ""},
      {"ftp://e.x/a b", "ftp://e.x/a b", ""},
      {"ftp://e.x/b#frag", "ftp://e.x/b", "frag"},
      {"news:x", "news:x", ""},
      {"HTTP://E.X:8080", "http://e.x:8080/", ""},
      {"https://user@E.X/", "https://user@E.X/", ""},
      {"HTTPS://User:Pw@E.X:443/P", "HTTPS://User:Pw@E.X:443/P", ""},
      {"https://e.x/%41%2f", "https://e.x/A%2f", ""},
  };
  for (const AddressRow& row : rows) {
    const NormalizedHyperlinkTarget got = normalize_hyperlink_target(row.input);
    EXPECT_EQ(got.target, row.target) << row.input;
    EXPECT_EQ(got.fragment, row.fragment) << row.input;
  }
}

TEST(HyperlinkAddress, EmptyInputStaysEmpty) {
  const NormalizedHyperlinkTarget got = normalize_hyperlink_target("");
  EXPECT_TRUE(got.target.empty());
  EXPECT_TRUE(got.fragment.empty());
}

}  // namespace
}  // namespace formulon
