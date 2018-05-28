---
title: ??????????? ??????? ?????????? ? ???????? ??? ????????? Office
description: ?????????????? ?????? ?????????? ??? Excel, Word, Outlook, PowerPoint ? OneNote.
ms.date: 03/23/2018
ms.openlocfilehash: f50ab7e5312702eb25fbb2c8a25291c5ff5027a7
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/23/2018
---
# <a name="office-add-in-host-and-platform-availability"></a>??????????? ??????? ?????????? ? ???????? ??? ????????? Office

?????? ?????????? Office ????? ???????? ?? ???????? ?????????? Office, ?????? ??????????, ???????? ??? ?????? API. ? ???????? ???? ???????????? ???????? ? ????????? ?????????, ?????? ??????????, ??????? ???????????? ????????? API ? ??????????? ??????? ???????????? ????????? API, ??????? ? ????????? ????? ?????????????? ??? ???? ?????????? Office. 

?????? * (?????????) ? ?????? ??????? ?????????, ??? ????????? ????? ????????. ? ???????? ?????????? ??? Project ? Access ????? ???????????? ? ?????? [??????????? ?????? ???????????? ????????? ??? Office](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).  

> [!NOTE]
> ????? ?????? ??? ?????? Office 2016, ?????????????? ? ??????? MSI, ? 16.0.4266.1001. ??? ?????? ???????? ?????? ????? ???????????? ????????? ExcelApi 1.1, WordApi 1.1 ? ??????????? ?????? ???????????? ????????? API.

## <a name="excel"></a>Excel

<table style="width:80%">
  <tr>
    <th style="width:10%">?????????</th>
    <th style="width:10%">????? ??????????</th> 
    <th style="width:20%">?????? ???????????? ????????? API</th> 
    <th style="width:40%"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>??????????? ???????? API</b></a></th> 
  </tr>
  <tr>
    <td>Office Online</td>
    <td> - ??????? ?????<br>
        - ???????<br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">??????? ??????????</a>
    </td>
    <td>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>
        - BindingEvents<br>
        - DocumentEvents<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - CompressedFile<br>
        - ?????????<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2013 ??? Windows</td>
    <td>
        - ??????? ?????<br>
        - ???????</td>
    <td>  - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>
        - BindingEvents<br>
        - DocumentEvents<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - ?????????<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2016 ??? Windows</td>
    <td>- ??????? ?????<br>
        - ???????<br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">??????? ?????????</a></td>
    <td>- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>- BindingEvents<br>
        - DocumentEvents<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - ?????????<br>
        - TextCoercion</td> 
  </tr>
  <tr>
    <td>Office ??? iOS</td>
    <td>- ??????? ?????<br>
        - ???????</td>
    <td>- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>- BindingEvents<br>
        - DocumentEvents<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - ?????????<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2016 ??? Mac</td>
    <td>- ??????? ?????<br>
        - ???????<br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">??????? ?????????</a></td>
    <td>- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>- BindingEvents<br>
        - DocumentEvents<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
</table>

<br/>

## <a name="outlook"></a>Outlook

<table style="width:80%">
  <tr>
    <th>?????????</th>
    <th>????? ??????????</th> 
    <th>?????? ???????????? ????????? API</th> 
    <th><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>??????????? ???????? API</b></a></th> 
  </tr>
  <tr>
    <td>Office Online</td>
    <td> - ?????? ?????<br>
      - ???????? ????????? ?????<br>
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">??????? ?????????</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></td>
    <td>??????????</td>
  </tr>
  <tr>
    <td>Office 2013 ??? Windows</td>
    <td> - ?????? ?????<br>
      - ???????? ????????? ?????<br>
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">??????? ?????????</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></td>
    <td>??????????</td>
  </tr>
  <tr>
    <td>Office 2016 ??? Windows</td>
    <td> - ?????? ?????<br>
      - ???????? ????????? ?????<br>
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">??????? ?????????</a><br>
      - ??????</td>
    <td> - <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></td>
    <td>??????????</td> 
  </tr>
  <tr>
    <td>Office ??? iOS</td>
    <td> - ?????? ?????<br>
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">??????? ?????????</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></td>    
    <td>??????????</td>
  </tr>
  <tr>
    <td>Office 2016 ??? Mac</td>
    <td> - ?????? ?????<br>
      - ???????? ????????? ?????<br>
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">??????? ?????????</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></td>
    <td>??????????</td>
  </tr>
  <tr>
    <td>Office ??? Android</td>
    <td> - ?????? ?????<br>
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">??????? ?????????</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></td>
    <td>??????????</td>
  </tr>
</table>

<br/>

## <a name="word"></a>Word

<table style="width:80%">
  <tr>
    <th>?????????</th>
    <th>????? ??????????</th> 
    <th>?????? ???????????? ????????? API</th> 
    <th><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>??????????? ???????? API</b></a></th> 
  </tr> 
  </tr>
  <tr>
    <td>Office Online</td>
    <td> - ??????? ?????<br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">??????? ?????????</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a><br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a><br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a><br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - BindingEvents<br>
         - CustomXmlParts<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - DocumentEvents<br>
         - TextFile<br>
         - ImageCoercion<br>
         - ?????????<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2013 ??? Windows</td>
    <td> - ??????? ?????</td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlPart<br>
         - DocumentEvents<br>
         - ????<br>
         - HtmlCoercion<br>
         - ImageCoercion<br>
         - OoxmlCoercion<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextFile<br>
         - ?????????<br>
         - TextCoercion<br>
         - MatrixCoercion<br>
         - ???????? ???????</td>
  </tr>
  <tr>
    <td>Office 2016 ??? Windows</td>
    <td> - ??????? ?????<br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">??????? ?????????</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlPart<br>
         - DocumentEvents<br>
         - ????<br>
         - HtmlCoercion<br>
         - ImageCoercion<br>
         - OoxmlCoercion<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextFile<br>
         - ?????????<br>
         - TextCoercion<br>
         - MatrixCoercion<br>
         - ???????? ??????? </td> 
  </tr>
  <tr>
    <td>Office ??? iOS</td>
    <td> - ??????? ?????</td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a><br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a><br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a><br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlPart<br>
         - DocumentEvents<br>
         - ????<br>
         - HtmlCoercion<br>
         - ImageCoercion<br>
         - OoxmlCoercion<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextFile<br>
         - ?????????<br>
         - TextCoercion<br>
         - MatrixCoercion<br>
         - ???????? ??????? </td> 
  </tr>
  <tr>
    <td>Office 2016 ??? Mac</td>
    <td> - ??????? ?????<br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">??????? ?????????</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlPart<br>
         - DocumentEvents<br>
         - ????<br>
         - HtmlCoercion<br>
         - ImageCoercion<br>
         - OoxmlCoercion<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextFile<br>
         - ?????????<br>
         - TextCoercion<br>
         - MatrixCoercion<br>
         - ???????? ??????? </td> 
  </tr>
</table>

<br/>

## <a name="powerpoint"></a>PowerPoint

<table style="width:80%">
  <tr>
    <th>?????????</th>
    <th>????? ??????????</th> 
    <th>?????? ???????????? ????????? API</th> 
    <th><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>??????????? ???????? API</b></a></th> 
  </tr> 
  </tr>
  <tr>
    <td>Office Online</td>
    <td> - ???????<br>
         - ??????? ?????<br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">??????? ?????????</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - ????<br>
         - ?????????<br>
         - ?????????<br>
         - TextCoercion<br>
         - ImageCoercion</td>
  </tr>
  <tr>
    <td>Office 2013 ??? Windows</td>
    <td> - ???????<br>
         - ??????? ?????<br>
    </td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - ????<br>
         - ?????????<br>
         - ?????????<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2016 ??? Windows</td>
    <td> - ???????<br>
         - ??????? ?????<br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">??????? ?????????</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - ????<br>
         - ?????????<br>
         - ?????????<br>
         - TextCoercion<br>
         - ImageCoercion</td>
  </tr>
  <tr>
    <td>Office ??? iOS</td>
    <td> - ???????<br>
         - ??????? ?????</td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
     <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - ????<br>
         - ?????????<br>
         - ?????????<br>
         - TextCoercion<br>
         - ImageCoercion</td>
  </tr>
  <tr>
    <td>Office 2016 ??? Mac</td>
    <td> - ???????<br>
         - ??????? ?????<br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">??????? ?????????</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - ????<br>
         - ?????????<br>
         - ?????????<br>
         - TextCoercion<br>
         - ImageCoercion</td>
  </tr>
</table>

<br/>

## <a name="onenote"></a>OneNote

<table style="width:80%">
  <tr>
    <th>?????????</th>
    <th>????? ??????????</th> 
    <th>?????? ???????????? ????????? API</th> 
    <th><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>??????????? ???????? API</b></a></th> 
  </tr> 
  </tr>
  <tr>
    <td>Office Online</td>
    <td> - ???????<br>
         - ??????? ?????<br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">??????? ?????????</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a><br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - DocumentEvents<br>
         - ?????????<br>
         - TextCoercion<br>
         - HtmlCoercion<br>
         - ImageCoercion</td>
  </tr>
  <tr>
    <td>Office 2013 ??? Windows</td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
  </tr> 
  <tr>
    <td>Office 2016 ??? Windows</td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td> 
  </tr>
  <tr>
    <td>Office ??? iOS</td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
  </tr>
  <tr>
    <td>Office 2016 ??? Mac</td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
  </tr>
</table>

<br/>

\* = ????????? ????? ????????. 

## <a name="see-also"></a>??. ?????

- [????? ????????? ????????? Office](office-add-ins.md)
- [??????????? ?????? ???????????? ????????? API](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)
- [?????? ???????????? ????????? ??? ?????? ?????????](https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets)
- [??????? ?? API JavaScript ??? Office](https://dev.office.com/reference/add-ins/javascript-api-for-office)

