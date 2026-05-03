// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.

#include "eval/rng.h"

namespace formulon {
namespace eval {

std::mt19937_64& thread_local_rng() {
  thread_local std::mt19937_64 rng{std::random_device{}()};
  return rng;
}

}  // namespace eval
}  // namespace formulon
