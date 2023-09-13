// taskpane.js
Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, function (args) {
  var selectedText = Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
          var text = result.value;
          // Abra o modal personalizado aqui.
          OfficeExtension.ExtensionHelpers.displayDialog("https://richardoliveira.github.io/modal.html", { width: 400, height: 200 });
      }
  });
});