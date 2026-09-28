// De onde o dashboard lê a planilha.
// Opção A (padrão): o arquivo planilha.xlsx enviado para este mesmo repositório.
// Opção B: link de uma planilha do Google Sheets publicada na web como .xlsx
//          (Arquivo > Compartilhar > Publicar na Web > Planilha inteira > Microsoft Excel (.xlsx)).
//          Cole o link abaixo no lugar de "planilha.xlsx".
window.DASH_CONFIG = {
  DATA_URL: "https://docs.google.com/spreadsheets/d/e/2PACX-1vRIft2UVppnF5D6ZhvURaTB9OPc0w5fhry96o2l-7MdOt3aMM2PyCEMlm3qrIaCF9sUcShsVomqsUXr/pub?output=xlsx"
};
