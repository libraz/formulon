void f() {
  fm_rec_t a = FM_STRUCT_INIT(fm_rec_t);
  fm_rec_t b = FM_STRUCT_INIT(fm_rec_t);
  const fm_other_t c = FM_STRUCT_INIT(fm_other_t);
  fm_cell_rect_t r{};
  fm_cell_rect_t r2{1, 2};
  fm_rec_t d{1, 2};
  fm_rec_t e = {.a = 1};
  fm_rec_t arr[2]{};
  fm_rec_t arr2[2] = {{}, {}};
  g(fm_rec_t FM_STRUCT_INIT(fm_rec_t));
  g(fm_rec_t{3, 4});
  // fm_rec_t in_comment{};
  const char* s = "fm_rec_t in_string{};";
  fm_rec_t* p{};
  fm_rec_t ok = FM_STRUCT_INIT(fm_rec_t);
}
