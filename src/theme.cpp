#include "theme.h"

namespace formulon {
namespace {

struct ScriptFaces {
  const char* script;
  const char* major;
  const char* minor;
};

// `a:font` entries of the Office theme, in the order Mac Excel 365 writes
// them into a new workbook's theme part (identical across UI locales).
constexpr ScriptFaces kOfficeScriptFaces[] = {
    {"Jpan", "游ゴシック Light", "游ゴシック"},
    {"Hang", "맑은 고딕", "맑은 고딕"},
    {"Hans", "等线 Light", "等线"},
    {"Hant", "新細明體", "新細明體"},
    {"Arab", "Times New Roman", "Arial"},
    {"Hebr", "Times New Roman", "Arial"},
    {"Thai", "Tahoma", "Tahoma"},
    {"Ethi", "Nyala", "Nyala"},
    {"Beng", "Vrinda", "Vrinda"},
    {"Gujr", "Shruti", "Shruti"},
    {"Khmr", "MoolBoran", "DaunPenh"},
    {"Knda", "Tunga", "Tunga"},
    {"Guru", "Raavi", "Raavi"},
    {"Cans", "Euphemia", "Euphemia"},
    {"Cher", "Plantagenet Cherokee", "Plantagenet Cherokee"},
    {"Yiii", "Microsoft Yi Baiti", "Microsoft Yi Baiti"},
    {"Tibt", "Microsoft Himalaya", "Microsoft Himalaya"},
    {"Thaa", "MV Boli", "MV Boli"},
    {"Deva", "Mangal", "Mangal"},
    {"Telu", "Gautami", "Gautami"},
    {"Taml", "Latha", "Latha"},
    {"Syrc", "Estrangelo Edessa", "Estrangelo Edessa"},
    {"Orya", "Kalinga", "Kalinga"},
    {"Mlym", "Kartika", "Kartika"},
    {"Laoo", "DokChampa", "DokChampa"},
    {"Sinh", "Iskoola Pota", "Iskoola Pota"},
    {"Mong", "Mongolian Baiti", "Mongolian Baiti"},
    {"Viet", "Times New Roman", "Arial"},
    {"Uigh", "Microsoft Uighur", "Microsoft Uighur"},
    {"Geor", "Sylfaen", "Sylfaen"},
    {"Armn", "Arial", "Arial"},
    {"Bugi", "Leelawadee UI", "Leelawadee UI"},
    {"Bopo", "Microsoft JhengHei", "Microsoft JhengHei"},
    {"Java", "Javanese Text", "Javanese Text"},
    {"Lisu", "Segoe UI", "Segoe UI"},
    {"Mymr", "Myanmar Text", "Myanmar Text"},
    {"Nkoo", "Ebrima", "Ebrima"},
    {"Olck", "Nirmala UI", "Nirmala UI"},
    {"Osma", "Ebrima", "Ebrima"},
    {"Phag", "Phagspa", "Phagspa"},
    {"Syrn", "Estrangelo Edessa", "Estrangelo Edessa"},
    {"Syrj", "Estrangelo Edessa", "Estrangelo Edessa"},
    {"Syre", "Estrangelo Edessa", "Estrangelo Edessa"},
    {"Sora", "Nirmala UI", "Nirmala UI"},
    {"Tale", "Microsoft Tai Le", "Microsoft Tai Le"},
    {"Talu", "Microsoft New Tai Lue", "Microsoft New Tai Lue"},
    {"Tfng", "Ebrima", "Ebrima"},
};

Theme build_default_theme() {
  Theme theme;
  theme.colors = {0xFF000000U, 0xFFFFFFFFU, 0xFF0E2841U, 0xFFE8E8E8U, 0xFF156082U, 0xFFE97132U,
                  0xFF196B24U, 0xFF0F9ED5U, 0xFFA02B93U, 0xFF4EA72EU, 0xFF467886U, 0xFF96607DU};
  for (const ScriptFaces& faces : kOfficeScriptFaces) {
    theme.major_scripts.push_back({faces.script, faces.major});
    theme.minor_scripts.push_back({faces.script, faces.minor});
  }
  // The Jpan entry heads the table; `*_ea` reports it while `a:ea` is empty.
  theme.fonts = {"Aptos Display", theme.major_scripts.front().typeface, "Aptos Narrow",
                 theme.minor_scripts.front().typeface};
  return theme;
}

}  // namespace

const Theme& default_theme() {
  static const Theme kTheme = build_default_theme();
  return kTheme;
}

}  // namespace formulon
