#include "theme.h"

#include <cstddef>
#include <cstdint>

namespace formulon {
namespace {

struct ScriptFace {
  const char* script;
  const char* face;
};

// `a:font` entries of the Office theme in the order Mac Excel 365 writes them
// into a new workbook's theme part (identical across UI locales); the face
// listed is the major one.
constexpr ScriptFace kOfficeMajorFaces[kOfficeScriptCount] = {
    {"Jpan", "游ゴシック Light"},
    {"Hang", "맑은 고딕"},
    {"Hans", "等线 Light"},
    {"Hant", "新細明體"},
    {"Arab", "Times New Roman"},
    {"Hebr", "Times New Roman"},
    {"Thai", "Tahoma"},
    {"Ethi", "Nyala"},
    {"Beng", "Vrinda"},
    {"Gujr", "Shruti"},
    {"Khmr", "MoolBoran"},
    {"Knda", "Tunga"},
    {"Guru", "Raavi"},
    {"Cans", "Euphemia"},
    {"Cher", "Plantagenet Cherokee"},
    {"Yiii", "Microsoft Yi Baiti"},
    {"Tibt", "Microsoft Himalaya"},
    {"Thaa", "MV Boli"},
    {"Deva", "Mangal"},
    {"Telu", "Gautami"},
    {"Taml", "Latha"},
    {"Syrc", "Estrangelo Edessa"},
    {"Orya", "Kalinga"},
    {"Mlym", "Kartika"},
    {"Laoo", "DokChampa"},
    {"Sinh", "Iskoola Pota"},
    {"Mong", "Mongolian Baiti"},
    {"Viet", "Times New Roman"},
    {"Uigh", "Microsoft Uighur"},
    {"Geor", "Sylfaen"},
    {"Armn", "Arial"},
    {"Bugi", "Leelawadee UI"},
    {"Bopo", "Microsoft JhengHei"},
    {"Java", "Javanese Text"},
    {"Lisu", "Segoe UI"},
    {"Mymr", "Myanmar Text"},
    {"Nkoo", "Ebrima"},
    {"Olck", "Nirmala UI"},
    {"Osma", "Ebrima"},
    {"Phag", "Phagspa"},
    {"Syrn", "Estrangelo Edessa"},
    {"Syrj", "Estrangelo Edessa"},
    {"Syre", "Estrangelo Edessa"},
    {"Sora", "Nirmala UI"},
    {"Tale", "Microsoft Tai Le"},
    {"Talu", "Microsoft New Tai Lue"},
    {"Tfng", "Ebrima"},
};

// The minor face equals the major one except for these scripts.
struct MinorOverride {
  std::uint8_t index;  // into kOfficeMajorFaces
  const char* face;
};
constexpr MinorOverride kOfficeMinorOverrides[] = {
    {0, "游ゴシック"}, {2, "等线"}, {4, "Arial"}, {5, "Arial"}, {10, "DaunPenh"}, {27, "Arial"},
};

Theme build_default_theme() {
  Theme theme;
  theme.colors = {0xFF000000U, 0xFFFFFFFFU, 0xFF0E2841U, 0xFFE8E8E8U, 0xFF156082U, 0xFFE97132U,
                  0xFF196B24U, 0xFF0F9ED5U, 0xFFA02B93U, 0xFF4EA72EU, 0xFF467886U, 0xFF96607DU};
  // The Jpan entry heads the table; `*_ea` reports it while `a:ea` is empty.
  theme.fonts = {"Aptos Display", office_script_face(0, true), "Aptos Narrow", office_script_face(0, false)};
  return theme;
}

}  // namespace

const char* office_script(std::size_t index) {
  return kOfficeMajorFaces[index].script;
}

const char* office_script_face(std::size_t index, bool major) {
  if (!major) {
    for (const MinorOverride& entry : kOfficeMinorOverrides) {
      if (entry.index == index) {
        return entry.face;
      }
    }
  }
  return kOfficeMajorFaces[index].face;
}

const Theme& default_theme() {
  static const Theme kTheme = build_default_theme();
  return kTheme;
}

}  // namespace formulon
