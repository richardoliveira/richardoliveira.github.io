// modal.js
document.addEventListener("DOMContentLoaded", function () {
  var formatButton = document.getElementById("formatButton");
  var cancelButton = document.getElementById("cancelButton");

  formatButton.addEventListener("click", function () {
      // Lógica para formatar o texto selecionado aqui
      // Você pode usar a API do Word para isso.
      Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
              var selectedText = result.value;
              // Faça algo com o texto selecionado, como formatação.
          }
      });
  });

  cancelButton.addEventListener("click", function () {
      // Feche o modal quando o botão Cancelar for clicado.
      OfficeExtension.ExtensionHelpers.closeDialog();
  });
});
