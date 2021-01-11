![GitHub repo size](https://img.shields.io/github/repo-size/felipebacelo/CombineWorksheets?style=for-the-badge)
![GitHub language count](https://img.shields.io/github/languages/count/felipebacelo/CombineWorksheets?style=for-the-badge)
![GitHub forks](https://img.shields.io/github/forks/felipebacelo/CombineWorksheets?style=for-the-badge)
![Bitbucket open pull requests](https://img.shields.io/bitbucket/pr-raw/felipebacelo/CombineWorksheets?style=for-the-badge)
![Bitbucket open issues](https://img.shields.io/bitbucket/issues/felipebacelo/CombineWorksheets?style=for-the-badge)

# CombineWorksheets
CombineWorksheets - VBA Excel

Simples solução para combinar diferentes pastas de trabalho do Excel em um único arquivo.

### Desenvolvimento

Desenvolvido em Microsoft VBA Excel.
***
### Requisitos

* Habilitar Macros
* Habilitar Guia de Desenvolvedor

### Referências às Bibliotecas

* Visual Basic For Applications
* Microsoft Excel 16.0 Object Library
* OLE Automation
* Microsoft Office 16.0 Object Library
* Microsoft Forms 2.0 Object Library

### Compatibilidade

Esta solução foi desenvolvida no Excel 2019 (64 bits) e testada no Excel 2016 (64 bits). Sua compatibilidade é garantida para a versão 2016 e superiores. Sua utilização em versões anteriores pode ocasionar em não funcionamento da mesma.

### Usabilidade

Para utilizar esta solução o usuário deverá:

* Realizar o download do arquivo ZIP: __CombineWorksheets__.
* Abrir o arquivo __CombineWorksheets.xlsm__, ou importar através do VBA os arquivos __Módulo1.bas__ e __UserForm1.frm__.
***
### Passo a Passo

Ao abrir o formulário será exibida a seguinte tela:

![Image1](https://github.com/felipebacelo/CombineWorksheets/blob/main/Images/Image1.jpg)

Será necessário selecionar a _Pasta de Arquivos Excel_, onde estão localizados os arquivos a serem combinados e a _Pasta de Destino_, onde o arquivo gerado será salvo.

Após a realização de todo processo, aparecerá a mensagem de confirmação com as informações sobre o arquivo gerado:

![Image1](https://github.com/felipebacelo/CombineWorksheets/blob/main/Images/Image2.jpg)
***
### Exemplo de Função Utilizada

```vba
Public Function PickFolder() As String

    If Application.FileDialog(4).Show Then
        PickFolder = Application.FileDialog(4).SelectedItems(1)
    End If

End Function
```

***
### Licenças

_MIT License_
_Copyright   ©   2020 Felipe Bacelo Rodrigues_
