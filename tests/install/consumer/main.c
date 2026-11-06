#include <stdio.h>

#include <formulon/formulon_c.h>

int main(void) {
  unsigned long runtime = (unsigned long)fm_get_version_number();
  unsigned long header = (unsigned long)FM_VERSION_NUMBER;
  printf("fm_get_version_number=%lu FM_VERSION_NUMBER=%lu version=%s\n", runtime, header,
         fm_get_version());
  return runtime == header ? 0 : 1;
}
