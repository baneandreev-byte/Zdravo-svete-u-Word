Office.onReady(() => {
  document.getElementById("btn").onclick = async () => {
    await Word.run(async (context) => {
      const sel = context.document.getSelection();
      sel.insertText("Zdravo svete", Word.InsertLocation.replace);
      await context.sync();
    });
  };
});
