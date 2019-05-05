function flatten(nestArrays) {
  return [].concat.apply([], nestArrays);
}


flatten([["a", "b"], "c"]);
