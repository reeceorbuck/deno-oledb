const denoDb = Deno.dlopen(
  "./oledb.cs",
  {
    "Invoke": {
      parameters: ["buffer"],
      result: "buffer",
      nonblocking: true,
    },
  } as const,
);

export default async (options: {dsn: string; query: string}) => {
  return await denoDb.symbols.Invoke(options)

  return new Promise(function (resolve, reject) {
    adodb(options, (error, result) => {
      if (error) {
        reject(new Error("deno-oledb failed", {cause: error}));
      } 
      resolve(result);
    });
  });
}