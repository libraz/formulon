#include "builtin_style_names.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>

#include "utils/number_text.h"

// @size-budget: 12 KB (name table, accent synthesizer and the cell-style save plan)

namespace formulon {
namespace {

constexpr std::uint32_t kAccentFirst = 29U;
constexpr std::uint32_t kAccentLast = 52U;
constexpr std::uint32_t kExplanatoryId = 53U;
constexpr std::uint32_t kStoredCount = 30U;

/// Per-locale template of the accent names "Accent n" and "p% - Accent n".
struct AccentTemplate {
  const char* word;      ///< Accent word of the plain form.
  const char* pct_word;  ///< Accent word after the percent prefix.
  std::uint8_t sep;      ///< Between word and n: 0 none, 1 space, 2 NBSP.
  std::uint8_t pct_sep;  ///< Between the percent number and "%": same coding.
  std::uint8_t dash;     ///< 0 " - ", 1 " <em dash> ".
};

constexpr const char* kSeps[3] = {"", " ", "\xC2\xA0"};
constexpr const char* kDashes[2] = {" - ", " \xE2\x80\x94 "};

// BEGIN GENERATED (tools/dev/gen_builtin_style_names.py)
// clang-format off
const char kNameBlob[] =
    "標準" "\0"
    "行レベル_" "\0"
    "列レベル_" "\0"
    "桁区切り [0.00]" "\0"
    "通貨 [0.00]" "\0"
    "パーセント" "\0"
    "桁区切り" "\0"
    "通貨" "\0"
    "ハイパーリンク" "\0"
    "表示済みのハイパーリンク" "\0"
    "メモ" "\0"
    "警告文" "\0"
    "強調 1" "\0"
    "強調 2" "\0"
    "強調 3" "\0"
    "タイトル" "\0"
    "見出し 1" "\0"
    "見出し 2" "\0"
    "見出し 3" "\0"
    "見出し 4" "\0"
    "入力" "\0"
    "出力" "\0"
    "計算" "\0"
    "チェック セル" "\0"
    "リンク セル" "\0"
    "集計" "\0"
    "良い" "\0"
    "悪い" "\0"
    "どちらでもない" "\0"
    "説明文" "\0"
    "Normal" "\0"
    "RowLevel_" "\0"
    "ColLevel_" "\0"
    "Comma" "\0"
    "Currency" "\0"
    "Percent" "\0"
    "Comma [0]" "\0"
    "Currency [0]" "\0"
    "Hyperlink" "\0"
    "Followed Hyperlink" "\0"
    "Note" "\0"
    "Warning Text" "\0"
    "Emphasis 1" "\0"
    "Emphasis 2" "\0"
    "Emphasis 3" "\0"
    "Title" "\0"
    "Heading 1" "\0"
    "Heading 2" "\0"
    "Heading 3" "\0"
    "Heading 4" "\0"
    "Input" "\0"
    "Output" "\0"
    "Calculation" "\0"
    "Check Cell" "\0"
    "Linked Cell" "\0"
    "Total" "\0"
    "Good" "\0"
    "Bad" "\0"
    "Neutral" "\0"
    "Explanatory Text" "\0"
    "Standard" "\0"
    "Zeilenebene_" "\0"
    "Spaltenebene_" "\0"
    "Komma" "\0"
    "Währung" "\0"
    "Prozent" "\0"
    "Dezimal [0]" "\0"
    "Währung [0]" "\0"
    "Link" "\0"
    "Besuchter Hyperlink" "\0"
    "Notiz" "\0"
    "Warnender Text" "\0"
    "Hervorhebung 1" "\0"
    "Hervorhebung 2" "\0"
    "Hervorhebung 3" "\0"
    "Überschrift" "\0"
    "Überschrift 1" "\0"
    "Überschrift 2" "\0"
    "Überschrift 3" "\0"
    "Überschrift 4" "\0"
    "Eingabe" "\0"
    "Ausgabe" "\0"
    "Berechnung" "\0"
    "Zelle überprüfen" "\0"
    "Verknüpfte Zelle" "\0"
    "Ergebnis" "\0"
    "Gut" "\0"
    "Schlecht" "\0"
    "Neutral" "\0"
    "Erklärender Text" "\0"
    "Normal" "\0"
    "NiveauLigne_" "\0"
    "NiveauColonne_" "\0"
    "Milliers" "\0"
    "Monétaire" "\0"
    "Pourcentage" "\0"
    "Milliers [0]" "\0"
    "Monétaire [0]" "\0"
    "Lien hypertexte" "\0"
    "Lien hypertexte visité" "\0"
    "Note" "\0"
    "Avertissement" "\0"
    "Accentuation 1" "\0"
    "Accentuation 2" "\0"
    "Accentuation 3" "\0"
    "Titre" "\0"
    "Titre 1" "\0"
    "Titre 2" "\0"
    "Titre 3" "\0"
    "Titre 4" "\0"
    "Entrée" "\0"
    "Sortie" "\0"
    "Calcul" "\0"
    "Vérification" "\0"
    "Cellule liée" "\0"
    "Total" "\0"
    "Satisfaisant" "\0"
    "Insatisfaisant" "\0"
    "Neutre" "\0"
    "Texte explicatif" "\0"
    "常规" "\0"
    "RowLevel_" "\0"
    "分级显示列_" "\0"
    "千位分隔" "\0"
    "货币" "\0"
    "百分比" "\0"
    "千位分隔[0]" "\0"
    "货币[0]" "\0"
    "超链接" "\0"
    "已访问的超链接" "\0"
    "注释" "\0"
    "警告文本" "\0"
    "强调 1" "\0"
    "强调 2" "\0"
    "强调 3" "\0"
    "标题" "\0"
    "标题 1" "\0"
    "标题 2" "\0"
    "标题 3" "\0"
    "标题 4" "\0"
    "输入" "\0"
    "输出" "\0"
    "计算" "\0"
    "检查单元格" "\0"
    "链接单元格" "\0"
    "汇总" "\0"
    "好" "\0"
    "差" "\0"
    "适中" "\0"
    "解释性文本" "\0"
    "표준" "\0"
    "RowLevel_" "\0"
    "ColLevel_" "\0"
    "쉼표" "\0"
    "통화" "\0"
    "백분율" "\0"
    "쉼표 [0]" "\0"
    "통화 [0]" "\0"
    "하이퍼링크" "\0"
    "열어 본 하이퍼링크" "\0"
    "메모" "\0"
    "경고문" "\0"
    "강조 1" "\0"
    "강조 2" "\0"
    "강조 3" "\0"
    "제목" "\0"
    "제목 1" "\0"
    "제목 2" "\0"
    "제목 3" "\0"
    "제목 4" "\0"
    "입력" "\0"
    "출력" "\0"
    "계산" "\0"
    "셀 확인" "\0"
    "연결된 셀" "\0"
    "요약" "\0"
    "좋음" "\0"
    "나쁨" "\0"
    "보통" "\0"
    "설명 텍스트" "\0"
    "ปกติ" "\0"
    "RowLevel_" "\0"
    "ColLevel_" "\0"
    "จุลภาค" "\0"
    "สกุลเงิน" "\0"
    "เปอร์เซ็นต์" "\0"
    "จุลภาค [0]" "\0"
    "สกุลเงิน [0]" "\0"
    "Hyperlink" "\0"
    "Followed Hyperlink" "\0"
    "หมายเหตุ" "\0"
    "ข้อความเตือน" "\0"
    "เน้นข้อความ 1" "\0"
    "เน้นข้อความ 2" "\0"
    "เน้นข้อความ 3" "\0"
    "ชื่อเรื่อง" "\0"
    "หัวเรื่อง 1" "\0"
    "หัวเรื่อง 2" "\0"
    "หัวเรื่อง 3" "\0"
    "หัวเรื่อง 4" "\0"
    "ป้อนค่า" "\0"
    "แสดงผล" "\0"
    "การคำนวณ" "\0"
    "เซลล์ตรวจสอบ" "\0"
    "เซลล์ที่มีลิงก์" "\0"
    "ผลรวม" "\0"
    "ดี" "\0"
    "แย่" "\0"
    "ปานกลาง" "\0"
    "ข้อความอธิบาย" "\0"
    "Normal" "\0"
    "NivelFila_" "\0"
    "NivelCol_" "\0"
    "Millares" "\0"
    "Moneda" "\0"
    "Porcentaje" "\0"
    "Millares [0]" "\0"
    "Moneda [0]" "\0"
    "Hipervínculo" "\0"
    "Hipervínculo visitado" "\0"
    "Notas" "\0"
    "Texto de advertencia" "\0"
    "Énfasis 1" "\0"
    "Énfasis 2" "\0"
    "Énfasis 3" "\0"
    "Título" "\0"
    "Encabezado 1" "\0"
    "Título 2" "\0"
    "Título 3" "\0"
    "Encabezado 4" "\0"
    "Entrada" "\0"
    "Salida" "\0"
    "Cálculo" "\0"
    "Celda de comprobación" "\0"
    "Celda vinculada" "\0"
    "Total" "\0"
    "Bueno" "\0"
    "Incorrecto" "\0"
    "Neutral" "\0"
    "Texto explicativo" "\0"
    "Normal" "\0"
    "NívelLinha_" "\0"
    "NívelCol_" "\0"
    "Vírgula" "\0"
    "Moeda" "\0"
    "Porcentagem" "\0"
    "Separador de milhares [0]" "\0"
    "Moeda [0]" "\0"
    "Hiperlink" "\0"
    "Hiperlink Visitado" "\0"
    "Nota" "\0"
    "Texto de Aviso" "\0"
    "Ênfase 1" "\0"
    "Ênfase 2" "\0"
    "Ênfase 3" "\0"
    "Título" "\0"
    "Título 1" "\0"
    "Título 2" "\0"
    "Título 3" "\0"
    "Título 4" "\0"
    "Entrada" "\0"
    "Saída" "\0"
    "Cálculo" "\0"
    "Célula de Verificação" "\0"
    "Célula Vinculada" "\0"
    "Total" "\0"
    "Bom" "\0"
    "Ruim" "\0"
    "Neutro" "\0"
    "Texto Explicativo" "\0"
    "Обычный" "\0"
    "УровеньСтрок_" "\0"
    "УровеньСтолб_" "\0"
    "Финансовый" "\0"
    "Денежный" "\0"
    "Процентный" "\0"
    "Финансовый [0]" "\0"
    "Денежный [0]" "\0"
    "Гиперссылка" "\0"
    "Открывавшаяся гиперссылка" "\0"
    "Примечание" "\0"
    "Текст предупреждения" "\0"
    "Выделение 1" "\0"
    "Выделение 2" "\0"
    "Выделение 3" "\0"
    "Название" "\0"
    "Заголовок 1" "\0"
    "Заголовок 2" "\0"
    "Заголовок 3" "\0"
    "Заголовок 4" "\0"
    "Ввод " "\0"
    "Вывод" "\0"
    "Вычисление" "\0"
    "Контрольная ячейка" "\0"
    "Связанная ячейка" "\0"
    "Итог" "\0"
    "Хороший" "\0"
    "Плохой" "\0"
    "Нейтральный" "\0"
    "Пояснение" "\0"
    "一般" "\0"
    "大綱列_" "\0"
    "大綱欄_" "\0"
    "千分位" "\0"
    "貨幣" "\0"
    "百分比" "\0"
    "千分位[0]" "\0"
    "貨幣 [0]" "\0"
    "超連結" "\0"
    "已瀏覽過的超連結" "\0"
    "備註" "\0"
    "警告文字" "\0"
    "強調 1" "\0"
    "強調 2" "\0"
    "強調 3" "\0"
    "標題" "\0"
    "標題 1" "\0"
    "標題 2" "\0"
    "標題 3" "\0"
    "標題 4" "\0"
    "輸入" "\0"
    "輸出" "\0"
    "計算方式" "\0"
    "檢查儲存格" "\0"
    "連結的儲存格" "\0"
    "合計" "\0"
    "好" "\0"
    "壞" "\0"
    "中等" "\0"
    "說明文字" "\0"
    "Normale" "\0"
    "LivelloRiga_" "\0"
    "LivelloCol_" "\0"
    "Migliaia" "\0"
    "Valuta" "\0"
    "Percentuale" "\0"
    "Migliaia [0]" "\0"
    "Valuta [0]" "\0"
    "Collegamento ipertestuale" "\0"
    "Collegamento ipertestuale visitato" "\0"
    "Nota" "\0"
    "Testo avviso" "\0"
    "Enfasi 1" "\0"
    "Enfasi 2" "\0"
    "Enfasi 3" "\0"
    "Titolo" "\0"
    "Titolo 1" "\0"
    "Titolo 2" "\0"
    "Titolo 3" "\0"
    "Titolo 4" "\0"
    "Input" "\0"
    "Output" "\0"
    "Calcolo" "\0"
    "Cella da controllare" "\0"
    "Cella collegata" "\0"
    "Totale" "\0"
    "Valore valido" "\0"
    "Valore non valido" "\0"
    "Neutrale" "\0"
    "Testo descrittivo" "\0"
    "Standaard" "\0"
    "RijNiveau_" "\0"
    "KolNiveau_" "\0"
    "Komma" "\0"
    "Valuta" "\0"
    "Procent" "\0"
    "Komma [0]" "\0"
    "Valuta [0]" "\0"
    "Hyperlink" "\0"
    "Gevolgde hyperlink" "\0"
    "Notitie" "\0"
    "Waarschuwingstekst" "\0"
    "Nadruk 1" "\0"
    "Nadruk 2" "\0"
    "Nadruk 3" "\0"
    "Titel" "\0"
    "Kop 1" "\0"
    "Kop 2" "\0"
    "Kop 3" "\0"
    "Kop 4" "\0"
    "Invoer" "\0"
    "Uitvoer" "\0"
    "Berekening" "\0"
    "Controlecel" "\0"
    "Gekoppelde cel" "\0"
    "Totaal" "\0"
    "Goed" "\0"
    "Ongeldig" "\0"
    "Neutraal" "\0"
    "Verklarende tekst" "\0";

const std::uint16_t kNameOffsets[13][31] = {
    {0, 7, 21, 35, 55, 69, 85, 98, 105, 127, 164, 171, 181, 190, 199, 208, 221, 233, 245, 257, 269, 276, 283, 290, 310, 327, 334, 341, 348, 370, 380},
    {380, 387, 397, 407, 413, 422, 430, 440, 453, 463, 482, 487, 500, 511, 522, 533, 539, 549, 559, 569, 579, 585, 592, 604, 615, 627, 633, 638, 642, 650, 667},
    {667, 676, 689, 703, 709, 718, 726, 738, 751, 756, 776, 782, 797, 812, 827, 842, 855, 870, 885, 900, 915, 923, 931, 942, 961, 979, 988, 992, 1001, 1009, 1027},
    {1027, 1034, 1047, 1062, 1071, 1082, 1094, 1107, 1122, 1138, 1162, 1167, 1181, 1197, 1213, 1229, 1235, 1244, 1253, 1262, 1271, 1279, 1286, 1293, 1307, 1321, 1327, 1340, 1355, 1362, 1379},
    {1379, 1386, 1396, 1413, 1426, 1433, 1443, 1459, 1469, 1479, 1501, 1508, 1521, 1530, 1539, 1548, 1555, 1564, 1573, 1582, 1591, 1598, 1605, 1612, 1628, 1644, 1651, 1655, 1659, 1666, 1682},
    {1682, 1689, 1699, 1709, 1716, 1723, 1733, 1744, 1755, 1771, 1798, 1805, 1815, 1824, 1833, 1842, 1849, 1858, 1867, 1876, 1885, 1892, 1899, 1906, 1917, 1931, 1938, 1945, 1952, 1959, 1976},
    {1976, 1989, 1999, 2009, 2028, 2053, 2087, 2110, 2139, 2149, 2168, 2193, 2230, 2266, 2302, 2338, 2369, 2399, 2429, 2459, 2489, 2511, 2530, 2555, 2592, 2638, 2654, 2661, 2671, 2693, 2733},
    {2733, 2740, 2751, 2761, 2770, 2777, 2788, 2801, 2812, 2826, 2849, 2855, 2876, 2887, 2898, 2909, 2917, 2930, 2940, 2950, 2963, 2971, 2978, 2987, 3010, 3026, 3032, 3038, 3049, 3057, 3075},
    {3075, 3082, 3095, 3106, 3115, 3121, 3133, 3159, 3169, 3179, 3198, 3203, 3218, 3228, 3238, 3248, 3256, 3266, 3276, 3286, 3296, 3304, 3311, 3320, 3345, 3363, 3369, 3373, 3378, 3385, 3403},
    {3403, 3418, 3444, 3470, 3491, 3508, 3529, 3554, 3575, 3598, 3648, 3669, 3709, 3730, 3751, 3772, 3789, 3810, 3831, 3852, 3873, 3883, 3894, 3915, 3951, 3983, 3992, 4007, 4020, 4043, 4062},
    {4062, 4069, 4080, 4091, 4101, 4108, 4118, 4131, 4142, 4152, 4177, 4184, 4197, 4206, 4215, 4224, 4231, 4240, 4249, 4258, 4267, 4274, 4281, 4294, 4310, 4329, 4336, 4340, 4344, 4351, 4364},
    {4364, 4372, 4385, 4397, 4406, 4413, 4425, 4438, 4449, 4475, 4510, 4515, 4528, 4537, 4546, 4555, 4562, 4571, 4580, 4589, 4598, 4604, 4611, 4619, 4640, 4656, 4663, 4677, 4695, 4704, 4722},
    {4722, 4732, 4743, 4754, 4760, 4767, 4775, 4785, 4796, 4806, 4825, 4833, 4852, 4861, 4870, 4879, 4885, 4891, 4897, 4903, 4909, 4916, 4924, 4935, 4947, 4962, 4969, 4974, 4983, 4992, 5010},
};

const AccentTemplate kAccentTemplates[13] = {
    {"アクセント", "アクセント", 1, 0, 0},
    {"Accent", "Accent", 0, 0, 0},
    {"Akzent", "Akzent", 0, 1, 0},
    {"Accent", "Accent", 0, 2, 0},
    {"着色", "着色", 1, 0, 0},
    {"강조색", "강조색", 0, 0, 0},
    {"ส่วนที่ถูกเน้น", "ส่วนที่ถูกเน้น", 0, 0, 0},
    {"Énfasis", "Énfasis", 0, 0, 0},
    {"Ênfase", "Ênfase", 0, 0, 0},
    {"Акцент", "акцент", 0, 0, 1},
    {"輔色", "輔色", 0, 0, 0},
    {"Colore", "Colore", 1, 0, 0},
    {"Accent", "Accent", 0, 0, 0},
};

// ExcelLocale -> row; locales with identical rows share one.
const std::uint8_t kLocaleRow[14] = {0, 1, 2, 3, 4, 5, 6, 7, 7, 8, 9, 10, 11, 12};
// clang-format on
// END GENERATED

void append_unsigned(std::string& out, std::uint32_t v) {
  char buf[16];
  const int n = format_unsigned(buf, sizeof(buf), v);
  out.append(buf, static_cast<std::size_t>(n));
}

std::string accent_name(std::uint32_t id, ExcelLocale locale) {
  const AccentTemplate& t = kAccentTemplates[kLocaleRow[static_cast<std::size_t>(locale)]];
  const std::uint32_t idx = id - kAccentFirst;
  const std::uint32_t k = idx % 4U;
  std::string out;
  if (k != 0U) {
    append_unsigned(out, 20U * k);
    out += kSeps[t.pct_sep];
    out += '%';
    out += kDashes[t.dash];
  }
  out += k != 0U ? t.pct_word : t.word;
  out += kSeps[t.sep];
  append_unsigned(out, idx / 4U + 1U);
  return out;
}

/// Stored name (or the stem of ids 1 and 2) of a non-accent id; empty for accent ids.
std::string_view builtin_style_base_name(std::uint32_t builtin_id, ExcelLocale locale) {
  if (builtin_id >= kAccentFirst && builtin_id != kExplanatoryId) {
    return {};
  }
  const std::size_t slot = builtin_id == kExplanatoryId ? kStoredCount - 1U : builtin_id;
  const std::uint16_t* offs = kNameOffsets[kLocaleRow[static_cast<std::size_t>(locale)]];
  return std::string_view(kNameBlob + offs[slot], static_cast<std::size_t>(offs[slot + 1U] - offs[slot]) - 1U);
}

}  // namespace

std::string builtin_style_name(std::uint32_t builtin_id, std::uint32_t i_level, ExcelLocale locale) {
  if (builtin_id >= kBuiltinStyleCount) {
    return {};
  }
  if (builtin_id >= kAccentFirst && builtin_id <= kAccentLast) {
    return accent_name(builtin_id, locale);
  }
  std::string out(builtin_style_base_name(builtin_id, locale));
  if (builtin_id == 1U || builtin_id == 2U) {
    append_unsigned(out, i_level + 1U);
  }
  return out;
}

}  // namespace formulon
