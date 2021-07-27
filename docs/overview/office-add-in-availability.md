---
title: Доступность клиентских приложений и платформ Office для надстроек Office
description: Поддерживаемые наборы обязательных элементов для Excel, OneNote, Outlook, PowerPoint, Project и Word.
ms.date: 07/13/2021
localization_priority: Priority
ms.openlocfilehash: 7b3bd770d74f29d1a0b27da5080284aa62146101
ms.sourcegitcommit: 30a861ece18255e342725e31c47f01960b854532
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/16/2021
ms.locfileid: "53455497"
---
# <a name="office-client-application-and-platform-availability-for-office-add-ins"></a>Доступность клиентских приложений и платформ Office для надстроек Office

Работа надстройки Office в соответствии с ожиданиями может зависеть от ведущего приложения Office, набора требований, элемента API или версии API. В таблицах ниже представлены сведения о доступных платформах, точках расширения, наборах обязательных элементов API и общих API, поддерживаемых в настоящее время для всех приложений Office.

<br>

|<a href="#excel"><img src="../images/index/logo-excel.svg" alt="Excel" width="48" /><br><span>Excel</span></a>|<a href="#onenote"><img src="../images/index/logo-onenote.svg" alt="OneNote" width="48" /><br><span>OneNote</span></a>|<a href="#outlook"><img src="../images/index/logo-outlook.svg" alt="Outlook" width="48" /><br><span>Outlook</span></a>|<a href="#powerpoint"><img src="../images/index/logo-powerpoint.svg" alt="PowerPoint" width="48" /><br><span>PowerPoint</span></a>|<a href="#project"><img src="../images/index/logo-project-server.svg" alt="Project" width="48" /><br><span>Project</span></a>|<a href="#word"><img src="../images/index/logo-word.svg" alt="Word" width="48" /><br><span>Word</span></a>|
|:---:|:---:|:---:|:---:|:---:|:---:|

> [!NOTE]
> Исходный выпуск Office 2016, установленный с помощью MSI, содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и наборы обязательных элементов общего API. Дополнительные сведения о журналах обновлений различных версий Office см. в разделе [См. также](#see-also). Надстройки Office могут не поддерживаться во всех службах, которые являются участниками [партнерской программы Office Cloud Storage](https://developer.microsoft.com/office/cloud-storage-partner-program), которая позволяет интегрировать Office в Интернете для работы с документами Office в рамках своего предложения услуг. Для получения дополнительных сведений обратитесь в службу-участника.

## <a name="excel"></a>Excel

<table style="width:80%">
  <tr>
    <th style="width:10%">Платформа</th>
    <th style="width:10%">Точки расширения</th>
    <th style="width:20%">Наборы обязательных элементов API</th>
    <th style="width:40%"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Общие API</b></a></th>
  </tr>
  <tr>
    <td>Office в Интернете</td>
    <td>
      - Область задач<br>
      - Контент<br>
      - CustomFunctions<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Команды надстройки</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a><br>
      - <a href="../reference/requirement-sets/excel-api-online-requirement-set.md">ExcelApiOnline</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/ribbon-api-requirement-sets.md">RibbonApi 1.1</a><br>
      - <a href="../reference/requirement-sets/shared-runtime-requirement-sets.md">SharedRuntime 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office для Windows<br>(подключено к подписке на Microsoft 365)</td>
    <td>
      - Область задач<br>
      - Контент<br>
      - CustomFunctions<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Команды надстройки</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a><br>
      - <a href="../reference/requirement-sets/ribbon-api-requirement-sets.md">RibbonApi 1.1</a><br>
      - <a href="../reference/requirement-sets/shared-runtime-requirement-sets.md">SharedRuntime 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2019 для Windows<br>(единовременная покупка)</td>
    <td>
      - Область задач<br>
      - Контент<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Команды надстройки</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2016 для Windows<br>(единовременная покупка)</td>
    <td>
      - Область задач<br>
      - Контент </td>
    <td>
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2013 для Windows<br>(единовременная покупка)</td>
    <td>
      - Область задач<br>
      - Контент </td>
    <td>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office для iPad<br>(подключено к подписке на Microsoft 365)</td>
    <td>
      - Область задач<br>
      - Контент </td>
    <td>
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office для Mac<br>(подключено к подписке на Microsoft 365)</td>
    <td>
      - Область задач<br>
      - Контент<br>
      - CustomFunctions<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Команды надстройки</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a><br>
      - <a href="../reference/requirement-sets/ribbon-api-requirement-sets.md">RibbonApi 1.1</a><br>
      - <a href="../reference/requirement-sets/shared-runtime-requirement-sets.md">SharedRuntime 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2019 для Mac<br>(единовременная покупка)</td>
    <td>
      - Область задач<br>
      - Контент<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Команды надстройки</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2016 для Mac<br>(единовременная покупка)</td>
    <td>
      - Область задач<br>
      - Контент </td>
    <td>
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
</table>

*&ast; - Добавлены обновления после выпуска.*

## <a name="custom-functions-excel-only"></a>Пользовательские функции (только Excel)

<table style="width:80%">
  <tr>
    <th>Платформа</th>
    <th>Точки расширения</th>
    <th>Наборы обязательных элементов API</th>
    <th><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Общие API</b></a></th>
  </tr>
  <tr>
    <td>Office в Интернете</td>
    <td>- CustomFunctions</td>
    <td>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.1</a><br>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.2</a><br>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.3</a>
    </td>
    <td></td>
  </tr>
  <tr>
    <td>Office для Windows<br>(подключено к подписке на Microsoft 365)</td>
    <td>- CustomFunctions</td>
    <td>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.1</a><br>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.2</a><br>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.3</a>
    </td>
    <td></td>
  </tr>
  <tr>
    <td>Office для Mac<br>(подключено к подписке на Microsoft 365)</td>
    <td>- CustomFunctions</td>
    <td>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.1</a><br>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.2</a><br>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.3</a>
    </td>
    <td></td>
  </tr>
</table>

## <a name="outlook"></a>Outlook

<table style="width:80%">
  <tr>
    <th>Платформа</th>
    <th>Точки расширения</th>
    <th>Наборы обязательных элементов API</th>
    <th><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Общие API</b></a></th>
  </tr>
  <tr>
    <td>Office в Интернете<br>(современная версия)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Чтение сообщения</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Создание сообщения</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Участник встречи (чтение)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Организатор встречи (создание)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Команды надстройки</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Mailbox 1.8</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md">Mailbox 1.9</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md">Mailbox 1.10</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup>
    </td>
    <td>Недоступно</td>
  </tr>
  <tr>
    <td>Office в Интернете<br>(классическая версия)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Чтение сообщения</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Создание сообщения</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Участник встречи (чтение)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Организатор встречи (создание)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Команды надстройки</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a>
    </td>
    <td>Недоступно</td>
  </tr>
  <tr>
    <td>Office для Windows<br>(подключено к подписке на Microsoft 365)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Чтение сообщения</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Создание сообщения</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Участник встречи (чтение)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Организатор встречи (создание)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Команды надстроек</a><br>
      - <a href="../reference/manifest/extensionpoint.md#module">Модули</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Mailbox 1.8</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md">Mailbox 1.9</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md">Mailbox 1.10</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>Недоступно</td>
  </tr>
  <tr>
    <td>Office 2019 для Windows<br>(единовременная покупка)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Чтение сообщения</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Создание сообщения</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Участник встречи (чтение)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Организатор встречи (создание)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Команды надстроек</a><br>
      - <a href="../reference/manifest/extensionpoint.md#module">Модули</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a>
    </td>
    <td>Недоступно</td>
  </tr>
  <tr>
    <td>Office 2016 для Windows<br>(единовременная покупка)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Чтение сообщения</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Создание сообщения</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Участник встречи (чтение)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Организатор встречи (создание)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Команды надстроек</a><br>
      - <a href="../reference/manifest/extensionpoint.md#module">Модули</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><sup>2</sup>
    </td>
    <td>Недоступно</td>
  </tr>
  <tr>
    <td>Office 2013 для Windows<br>(единовременная покупка)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Чтение сообщения</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Создание сообщения</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Участник встречи (чтение)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Организатор встречи (создание)</a><br>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><sup>2</sup><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><sup>2</sup>
    </td>
    <td>Недоступно</td>
  </tr>
  <tr>
    <td>Office для iOS<br>(подключено к подписке на Microsoft 365)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface">Чтение сообщения</a><br>
      - <a href="../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface">Организатор встречи (создание): собрание по сети</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Команды надстройки</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a>
    </td>
    <td>Недоступно</td>
  </tr>
  <tr>
    <td>Office для Mac<br>(текущий пользовательский интерфейс,<br>подключенный к подписке на Microsoft 365)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Чтение сообщения</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Создание сообщения</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Участник встречи (чтение)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Организатор встречи (создание)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Команды надстройки</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Mailbox 1.8</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>Недоступно</td>
  </tr>
  <tr>
    <td>Office для Mac<br>(новый пользовательский интерфейс (предварительная версия)<sup>3</sup>,<br>подключенный к подписке на Microsoft 365)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Чтение сообщения</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Создание сообщения</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Участник встречи (чтение)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Организатор встречи (создание)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Команды надстройки</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Mailbox 1.8</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup>
    </td>
    <td>Недоступно</td>
  </tr>
  <tr>
    <td>Office 2019 для Mac<br>(единовременная покупка)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Чтение сообщения</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Создание сообщения</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Участник встречи (чтение)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Организатор встречи (создание)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Команды надстройки</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a>
    </td>
    <td>Недоступно</td>
  </tr>
  <tr>
    <td>Office 2016 для Mac<br>(единовременная покупка)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Чтение сообщения</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Создание сообщения</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Участник встречи (чтение)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Организатор встречи (создание)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Команды надстройки</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a>
    </td>
    <td>Недоступно</td>
  </tr>
  <tr>
    <td>Office для Android<br>(подключено к подписке на Microsoft 365)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface">Чтение сообщения</a><br>
      - <a href="../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface">Организатор встречи (создание): собрание по сети</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Команды надстройки</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a>
    </td>
    <td>Недоступно</td>
  </tr>
</table>

> [!NOTE]
> <sup>1</sup>Чтобы потребовать набор API удостоверений 1.3 в коде надстройки, проверьте, поддерживается ли он вызовом `isSetSupported('IdentityAPI', '1.3')`. Объявление в манифесте надстройки не поддерживается. Также можно определить, поддерживается ли API, проверив, не `undefined` ли он. Подробнее см. в статье [Использование API из наборов требования более поздних версий](../reference/requirement-sets/outlook-api-requirement-sets.md#using-apis-from-later-requirement-sets).
>
> <sup>2</sup> Добавлены обновления после выпуска.
>
> <sup>3</sup> Поддержка нового пользовательского интерфейса Mac (предварительной версии) доступна в Outlook с версии 16.38.506. Дополнительные сведения см. в разделе [Поддержка надстроек в Outlook в новом интерфейсе Mac](../outlook/compare-outlook-add-in-support-in-outlook-for-mac.md#add-in-support-in-outlook-on-new-mac-ui-preview).

> [!IMPORTANT]
> Поддержка клиентами набора обязательных элементов может ограничиваться поддержкой сервера Exchange. Подробные сведения о диапазоне наборов обязательных элементов, поддерживаемых сервером Exchange и клиентами Outlook, см. в статье [Наборы обязательных элементов API JavaScript для Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).

<br/>

## <a name="word"></a>Word

<table style="width:80%">
  <tr>
    <th>Платформа</th>
    <th>Точки расширения</th>
    <th>Наборы обязательных элементов API</th>
    <th><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Общие API</b></a></th>
  </tr>
  <tr>
    <td>Office в Интернете</td>
    <td>
      - Область задач<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Команды надстройки</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a><br>
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a><br>
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Office для Windows<br>(подключено к подписке на Microsoft 365)</td>
    <td>
      - Область задач<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Команды надстройки</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a><br>
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a><br>
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Office 2019 для Windows<br>(единовременная покупка)</td>
    <td>
      - Область задач<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Команды надстройки</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a><br>
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a><br>
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Office 2016 для Windows<br>(единовременная покупка)</td>
    <td>- Область задач</td>
    <td>
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Office 2013 для Windows<br>(единовременная покупка)</td>
    <td>- Область задач</td>
    <td>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Office для iPad<br>(подключено к подписке на Microsoft 365)</td>
    <td>- Область задач</td>
    <td>
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a><br>
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a><br>
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Office для Mac<br>(подключено к подписке на Microsoft 365)</td>
    <td>
      - Область задач<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Команды надстройки</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a><br>
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a><br>
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Office 2019 для Mac<br>(единовременная покупка)</td>
    <td>
      - Область задач<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Команды надстройки</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a><br>
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a><br>
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Office 2016 для Mac<br>(единовременная покупка)</td>
    <td>- Область задач</td>
    <td>
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
</table>

*&ast; - Добавлены обновления после выпуска.*

<br/>

## <a name="powerpoint"></a>PowerPoint

<table style="width:80%">
  <tr>
    <th>Платформа</th>
    <th>Точки расширения</th>
    <th>Наборы обязательных элементов API</th>
    <th><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Общие API</b></a></th>
  </tr>
  <tr>
    <td>Office в Интернете</td>
    <td>
      - Контент<br>
      - Область задач<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Команды надстройки</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a><br>
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office для Windows<br>(подключено к подписке на Microsoft 365)</td>
    <td>
      - Контент<br>
      - Область задач<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Команды надстройки</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a><br>
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2019 для Windows<br>(единовременная покупка)</td>
    <td>
      - Контент<br>
      - Область задач<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Команды надстройки</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2016 для Windows<br>(единовременная покупка)</td>
    <td>
      - Контент<br>
      - Область задач </td>
    <td>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2013 для Windows<br>(единовременная покупка)</td>
    <td>
      - Контент<br>
      - Область задач </td>
    <td>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office для iPad<br>(подключено к подписке на Microsoft 365)</td>
    <td>
      - Контент<br>
      - Область задач </td>
    <td>
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a><br>
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office для Mac<br>(подключено к подписке на Microsoft 365)</td>
    <td>
      - Контент<br>
      - Область задач<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Команды надстройки</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a><br>
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2019 для Mac<br>(единовременная покупка)</td>
    <td>
      - Контент<br>
      - Область задач<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Команды надстройки</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2016 для Mac<br>(единовременная покупка)</td>
    <td>
      - Контент<br>
      - Область задач </td>
    <td>
       - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
</table>

*&ast; - Добавлены обновления после выпуска.*

<br/>

## <a name="onenote"></a>OneNote

<table style="width:80%">
  <tr>
    <th>Платформа</th>
    <th>Точки расширения</th>
    <th>Наборы обязательных элементов API</th>
    <th><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Общие API</b></a></th>
  </tr>
  <tr>
    <td>Office в Интернете</td>
    <td>
      - Контент<br>
      - Область задач<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Команды надстройки</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/onenote-api-requirement-sets.md">OneNoteApi 1.1</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
</table>

<br/>

## <a name="project"></a>Project

<table style="width:80%">
  <tr>
    <th>Платформа</th>
    <th>Точки расширения</th>
    <th>Наборы обязательных элементов API</th>
    <th><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Общие API</b></a></th>
  </tr>
  <tr>
    <td>Office 2019 для Windows<br>(единовременная покупка)</td>
    <td>- Область задач</td>
    <td>- <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2016 для Windows<br>(единовременная покупка)</td>
    <td>- Область задач</td>
    <td>- <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2013 для Windows<br>(единовременная покупка)</td>
    <td>- Область задач</td>
    <td>- <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
</table>

<br/>

## <a name="see-also"></a>См. также

- [Обзор платформы надстроек Office](office-add-ins.md)
- [Версии Office и наборы обязательных элементов](../develop/office-versions-and-requirement-sets.md)
- [Наборы обязательных элементов общего API](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [Наборы обязательных элементов для команд надстроек](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [Справочная документация по API](../reference/javascript-api-for-office.md)
- [Журнал обновлений для Приложений Microsoft 365](/officeupdates/update-history-office365-proplus-by-date)
- [Журнал обновлений Office 2016 и 2019 ("нажми и работай")](/officeupdates/update-history-office-2019)
- [Журнал обновлений Office 2013 ("нажми и работай")](/officeupdates/update-history-office-2013)
- [Журнал обновлений Office 2010, 2013 и 2016 (MSI)](/officeupdates/office-updates-msi)
- [Журнал обновлений Outlook 2010, 2013 и 2016 (MSI)](/officeupdates/outlook-updates-msi)
- [Журнал обновлений Office для Mac](/officeupdates/update-history-office-for-mac)
- [Разработка надстроек Office](../develop/develop-overview.md)
