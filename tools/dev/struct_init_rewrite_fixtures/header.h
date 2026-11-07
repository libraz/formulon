typedef struct {
  uint32_t struct_size;  /* record */
  int a;
  int b;
} fm_rec_t;

typedef struct {
  // record with a comment first
  uint32_t struct_size;
  union { int x; } u;
} fm_other_t;

typedef struct {
  int row;
  int col;
} fm_cell_rect_t;  /* frozen value type */
