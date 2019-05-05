function flatten(nestArrays) {
  return [].concat.apply([], nestArrays);
}


flatten([["a", "b"], "c"]);


// let a;
// a = [].concat(["a", "b"]);
// a = a.concat("c");
// イメージ的には ↑ のように、[] に対して nestArrays の各要素を concat していってくれるので、flatt にできる
