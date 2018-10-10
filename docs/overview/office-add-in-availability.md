---
title: Доступность ведущих приложений и платформ для надстроек Office
description: Поддерживаемые наборы обязательных элементов  для Excel, Word, Outlook, PowerPoint и OneNote.
ms.date: 10/03/2018
ms.openlocfilehash: 6f7b5b565773457e6cd8a9eee69eb304784a29a9
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459317"
---
# <a name="office-add-in-host-and-platform-availability"></a>Доступность ведущих приложений и платформ для надстроек Office

Работа надстройки Office должным образом может зависеть от ведущего приложения Office, набора требований, элемента или версии API. В таблицах ниже представлены сведения о доступной платформе, точках расширения, наборах обязательных элементов API и стандартных наборах обязательных элементов API,  которые в настоящее время поддерживаются для всех приложений Office.

Если ячейка таблицы содержит символ звездочки (*), это означает, что поддержка скоро появится. С наборами требований для Project и Access можно ознакомиться в статье [Стандартные наборы обязательных элементов для Office](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets).  

> [!NOTE]
> Номер сборки для набора Office 2016, установленного с помощью MSI, — 16.0.4266.1001. Эта версия содержит только набор обязательных элементов ExcelApi 1.1, WordApi 1.1 и стандартные наборы обязательных элементов API.

## <a name="excel"></a>Excel

<table style="width:80%">
  <tr>
    <th style="width:10%">Платформа</th>
    <th style="width:10%">Точки расширения</th>
    <th style="width:20%">Наборы обязательных элементов API</th>
    <th style="width:40%"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></th>
  </tr>
  <tr>
    <td>Office Online</td>
    <td> - Панель задач<br>
        - Содержимое<br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a>
    </td>
    <td>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>
        - BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - File<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - Selection<br>
        - Settings<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2013 для Windows</td>
    <td>
        - Панель задач<br>
        - Содержимое</td>
    <td>  - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>
        - BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - File<br>
        - ImageCoercion<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - Selection<br>
        - Settings<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2016 для Windows</td>
    <td>- Панель задач<br>
        - Содержимое<br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></td>
    <td>- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>- BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - File<br>
        - ImageCoercion<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - Selection<br>
        - Settings<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2019 для Windows</td>
    <td>- Панель задач<br>
        - Содержимое<br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></td>
    <td>- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>- BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - File<br>
        - ImageCoercion<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - Selection<br>
        - Settings<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Office для iOS</td>
    <td>- Панель задач<br>
        - Содержимое</td>
    <td>- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a><br>
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>- BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - File<br>
        - ImageCoercion<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - Selection<br>
        - Settings<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2016 для Mac</td>
    <td>- Панель задач<br>
        - Содержимое<br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></td>
    <td>- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>- BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - File<br>
        - ImageCoercion<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - PdfFile<br>
        - Selection<br>
        - Settings<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2019 для Mac</td>
    <td>- Панель задач<br>
        - Содержимое<br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></td>
    <td>- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>- BindingEvents<br>
        - CompressedFile<br>
        - DocumentEvents<br>
        - File<br>
        - ImageCoercion<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - PdfFile<br>
        - Selection<br>
        - Settings<br>
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
    <th>Платформа</th>
    <th>Точки расширения</th>
    <th>Наборы обязательных элементов API</th>
    <th><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></th>
  </tr>
  <tr>
    <td>Office Online</td>
    <td> - Чтение почты<br>
      -  Создание сообщения почты<br>
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></td>
    <td>Недоступна</td>
  </tr>
  <tr>
    <td>Office 2013 для Windows</td>
    <td> - Чтение почты<br>
      -  Создание сообщения почты<br>
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></td>
    <td>Недоступна</td>
  </tr>
  <tr>
    <td>Office 2016 для Windows</td>
    <td> - Чтение почты<br>
      -  Создание сообщения почты<br>
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a><br>
      - Модули</td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></td>
    <td>Недоступна</td>
  </tr>
  <tr>
    <td>Office 2019 для Windows</td>
    <td> - Чтение почты<br>
      -  Создание сообщения почты<br>
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a><br>
      - Модули</td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></td>
    <td>Недоступна</td>
  </tr>
  <tr>
    <td>Office для iOS</td>
    <td> - Чтение почты<br>
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></td>
    <td>Недоступна</td>
  </tr>
  <tr>
    <td>Office 2016 для Mac</td>
    <td> - Чтение почты<br>
      -  Создание сообщения почты<br>
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></td>
    <td>Недоступна</td>
  </tr>
  <tr>
    <td>Office 2019 для Mac</td>
    <td> - Чтение почты<br>
      -  Создание сообщения почты<br>
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></td>
    <td>Недоступна</td>
  </tr>
  <tr>
    <td>Office для Android</td>
    <td> - Чтение почты<br>
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a><br>
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></td>
    <td>Недоступна</td>
  </tr>
</table>

<br/>

## <a name="word"></a>Word

<table style="width:80%">
  <tr>
    <th>Платформа</th>
    <th>Точки расширения</th>
    <th>Наборы обязательных элементов API</th>
    <th><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></th>
  </tr> 
  </tr>
  <tr>
    <td>Office Online</td>
    <td> - Панель задач<br>
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a><br>
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a><br>
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a><br>
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - BindingEvents<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - ImageCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile</td>
  </tr>
  <tr>
    <td>Office 2013 для Windows</td>
    <td> - Панель задач</td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - ImageCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile</td>
  </tr>
  <tr>
    <td>Office 2016 для Windows</td>
    <td> - Панель задач<br>
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - ImageCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile </td>
  </tr>
  <tr>
    <td>Office 2019 для Windows</td>
    <td> - Панель задач<br>
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - ImageCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile </td>
  </tr>
  <tr>
    <td>Office для iOS</td>
    <td> - Панель задач</td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a><br>
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a><br>
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a><br>
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - ImageCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile </td>
  </tr>
  <tr>
    <td>Office 2016 для Mac</td>
    <td> - Панель задач<br>
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - ImageCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile </td>
  </tr>
  <tr>
    <td>Office 2019 для Mac</td>
    <td> - Панель задач<br>
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a><br>
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlParts<br>
         - DocumentEvents<br>
         - File<br>
         - HtmlCoercion<br>
         - ImageCoercion<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - OoxmlCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextCoercion<br>
         - TextFile </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a>PowerPoint

<table style="width:80%">
  <tr>
    <th>Платформа</th>
    <th>Точки расширения</th>
    <th>Наборы обязательных элементов API</th>
    <th><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></th>
  </tr> 
  </tr>
  <tr>
    <td>Office Online</td>
    <td> - Содержимое<br>
         - Панель задач<br>
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - ImageCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2013 для Windows</td>
    <td> - Содержимое<br>
         - Панель задач<br>
    </td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - ImageCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2016 для Windows</td>
    <td> - Содержимое<br>
         - Панель задач<br>
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - ImageCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2019 для Windows</td>
    <td> - Содержимое<br>
         - Панель задач<br>
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - ImageCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Office для iOS</td>
    <td> - Содержимое<br>
         - Панель задач</td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
     <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion<br>
         - ImageCoercion</td>
  </tr>
  <tr>
    <td>Office 2016 для Mac</td>
    <td> - Содержимое<br>
         - Панель задач<br>
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - ImageCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2019 для Mac</td>
    <td> - Содержимое<br>
         - Панель задач<br>
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - File<br>
         - ImageCoercion<br>
         - PdfFile<br>
         - Selection<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
</table>

<br/>

## <a name="onenote"></a>OneNote

<table style="width:80%">
  <tr>
    <th>Платформа</th>
    <th>Точки расширения</th>
    <th>Наборы обязательных элементов API</th>
    <th><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Стандартные элементы API</b></a></th>
  </tr> 
  </tr>
  <tr>
    <td>Office Online</td>
    <td> - Содержимое<br>
         - Панель задач<br>
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Команды надстроек</a></td>
    <td> - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a><br>
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - DocumentEvents<br>
         - HtmlCoercion<br>
         - ImageCoercion<br>
         - Settings<br>
         - TextCoercion</td>
  </tr>
</table>

<br/>

## <a name="see-also"></a>См. также

- [Обзор платформы надстроек Office](office-add-ins.md)
- [Стандартные наборы обязательных элементов API](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets)
- [Наборы обязательных элементов для команд надстроек](https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets)
- [Ссылка на API JavaScript для Office](https://docs.microsoft.com/javascript/office/javascript-api-for-office)
