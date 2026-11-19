// The built-in defined names Excel knows and their spelling per locale.

#ifndef FORMULON_BUILTIN_DEFINED_NAMES_H_
#define FORMULON_BUILTIN_DEFINED_NAMES_H_

#include <string_view>

#include "utils/strings.h"

namespace formulon {

// A built-in defined name and its spelling per locale, in `ExcelLocale`
// order; an empty field is the English name. es_ES and es_MX agree.
struct BuiltinNameSpelling {
  std::string_view english;
  std::string_view local[14];
};

// ja en de fr zh_CN ko th es_ES es_MX pt_BR ru zh_TW it nl
inline constexpr BuiltinNameSpelling kBuiltinNameSpellings[] = {
    {"Print_Area",
     {"", "", "Druckbereich", "Zone_d_impression", "", "", "", "Área_de_impresión", "Área_de_impresión",
      "Area_de_impressao", "Область_печати", "", "Area_stampa", "Afdrukbereik"}},
    {"Print_Titles",
     {"", "", "Drucktitel", "Impression_des_titres", "", "", "", "Títulos_a_imprimir", "Títulos_a_imprimir",
      "Titulos_de_impressao", "Заголовки_для_печати", "", "Titoli_stampa", "Afdruktitels"}},
    {"_FilterDatabase",
     {"", "", "_FilterDatenbank", "", "", "", "", "", "", "_FiltrarBancodeDados", "_ФильтрБазыДанных", "",
      "_FiltroDatabase", ""}},
    {"Criteria",
     {"", "", "Suchkriterien", "Criteres", "", "", "", "Criterios", "Criterios", "Criterios", "Критерии", "", "Criteri",
      ""}},
    {"Extract",
     {"", "", "Zielbereich", "Extraire", "提取", "", "", "Área_de_extracción", "Área_de_extracción", "Area_de_extracao",
      "Извлечь", "", "Estrai", "Ophalen"}},
    {"Database",
     {"", "", "Datenbank", "Base_de_donnees", "", "", "", "BaseDeDatos", "BaseDeDatos", "Banco_de_dados", "База_данных",
      "", "", ""}},
    {"Consolidate_Area",
     {"", "", "Konsolidierungsbereich", "Zone_de_consolidation", "", "", "", "Área_de_consolidación",
      "Área_de_consolidación", "Area_de_consolidacao", "Область_консолидации", "", "Area_unifica",
      "Consolidatiebereik"}},
    {"Sheet_Title",
     {"", "", "Tabellenüberschrift", "Titre_de_la_feuille", "", "", "", "Título_hoja", "Título_hoja",
      "Título_de_planilha", "Название_листа", "", "Titolo_foglio", "Bladtitel"}},
    {"Auto_Open",
     {"", "", "Auto_öffnen", "Auto_ouvrir", "", "", "", "Auto_abrir", "Auto_abrir", "Abrir_auto", "Авто_открыть", "",
      "Auto_apri", "Auto_openen"}},
    {"Recorder",
     {"", "", "Aufzeichnung", "Fonction_enregistrement", "录音机", "", "", "Grabación", "Grabación", "Gravador",
      "Рекордер", "", "Registratore", "Opname"}},
    {"Auto_Close",
     {"", "", "Auto_schließen", "Auto_fermer", "", "", "", "Auto_cerrar", "Auto_cerrar", "Fechar_auto", "Авто_закрыть",
      "", "Auto_chiudi", "Auto_sluiten"}},
    {"Data_Form",
     {"", "", "Datenmaske", "Grille", "", "", "", "Ficha_de_datos", "Ficha_de_datos", "Formulario_de_dados",
      "Форма_данных", "", "Modulo_dati", "Gegevensformulier"}},
    {"Auto_Activate",
     {"", "", "Auto_aktivieren", "Auto_activer", "", "", "", "Auto_activar", "Auto_activar", "Ativar_auto",
      "Авто_активизировать", "", "Auto_attiva", "Auto_activeren"}},
    {"Auto_Deactivate",
     {"", "", "Auto_deaktivieren", "Auto_desactiver", "", "", "", "Auto_desactivar", "Auto_desactivar",
      "Desativar_auto", "Авто_дезактивизировать", "", "Auto_disattiva", "Auto_deactiveren"}},
};

/// The English spelling of the built-in defined name `name` (no `_xlnm.` prefix, ASCII case-insensitive), or empty
/// when it is none.
inline std::string_view builtin_name_english(std::string_view name) noexcept {
  for (const BuiltinNameSpelling& spelling : kBuiltinNameSpellings) {
    if (strings::case_insensitive_eq(spelling.english, name)) {
      return spelling.english;
    }
  }
  return {};
}

}  // namespace formulon

#endif  // FORMULON_BUILTIN_DEFINED_NAMES_H_
