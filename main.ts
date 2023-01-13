import edge from "npm:edge-js";

const adodb = edge.func("./oledb.cs");

export default (options: string) => {
  return new Promise(function (resolve, reject) {
    adodb(options, (error, result) => {
      if (error) {
        reject(new Error("deno-oledb failed", {cause: error}));
      } 
      resolve(result);
    });
  });
}