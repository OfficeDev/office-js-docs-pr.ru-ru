---
title: Объектная модель JavaScript для Word в надстройках Office
description: Узнайте о ключевых компонентах объектной модели JavaScript, определенной в Word.
ms.date: 3/17/2022
ms.localizationpriority: high
ms.openlocfilehash: d3c2a43e2febbf31fe132dfb5c220bffcc7a1fef
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746102"
---
# <a name="word-javascript-object-model-in-office-add-ins"></a>Объектная модель JavaScript для Word в надстройках Office

В этой статье описаны основные концепции использования [API JavaScript для Word](../reference/overview/word-add-ins-reference-overview.md) с целью создания надстроек.

> [!IMPORTANT]
> Сведения об асинхронном типе API-интерфейсов Word и принципах их работы с документами см. в статье [Использование модели API, зависящей от приложения](../develop/application-specific-api-model.md).

## <a name="officejs-apis-for-word"></a>API-интерфейсы Office.js для Word

Надстройка Word взаимодействует с объектами в Word с помощью API JavaScript для Office. Сюда относятся две объектные модели JavaScript:

* **API JavaScript для Word**. [API-интерфейс JavaScript для Word](/javascript/api/word) предоставляет строго типизированные объекты, подходящие для документов, диапазонов, таблиц, списков, форматирования и т. д.

* **Общие API-интерфейсы**. [Общий API](/javascript/api/office) предоставляет доступ к таким компонентам, как пользовательский интерфейс, диалоговые окна и параметры клиентов, общие для разных приложений Office.

Скорее всего, вы будете разрабатывать большую часть функций надстроек для Word с помощью API JavaScript для Word, но вам также потребуются объекты из общего API. Например:

* [Office.Context](/javascript/api/office/office.context). Объект `Context` представляет среду выполнения надстройки и предоставляет доступ к ключевым объектам API. Он состоит из данных о конфигурации, например `contentLanguage` и `officeTheme`, а также предоставляет сведения о среде выполнения надстройки, например `host` и `platform`. Кроме того, он предоставляет метод `requirements.isSetSupported()`, с помощью которого можно проверить, поддерживается ли указанный набор обязательных элементов приложением Word, в котором запускается надстройка.
* [Office.Document](/javascript/api/office/office.document). Объект `Office.Document` предоставляет метод `getFileAsync()`, позволяющий загрузить файл Word, в котором работает надстройка. Это выполняется отдельно от объекта [Word.Document](/javascript/api/word/word.document).

![Различия между API JS для Word и общими API.](../images/word-js-api-common-api.png)

## <a name="word-specific-object-model"></a>Объектная модель для Word

Чтобы понять API-интерфейсы Word, нужно понимать, как компоненты документа связаны друг с другом.

* Объект **Document** содержит объекты **Section**, а также объекты уровня документа, например параметры и настраиваемые части XML.
* Объект **Section** содержит объект **Body**.
* Объект **Body** предоставляет доступ к объектам **Paragraph**, **ContentControl** и **Range**, а также к другим объектам.
* Объект **Range** представляет собой непрерывную область содержимого, включающую текст, пробелы, объекты **Table**, а также изображения. Он также содержит большую часть методов обработки текста.
* Объект **List** представляет текст в виде нумерованного или маркированного списка.

## <a name="see-also"></a>См. также

- [Обзор API JavaScript для Word](../reference/overview/word-add-ins-reference-overview.md)
- [Создание первой надстройки Word](../quickstarts/word-quickstart.md)
- [Руководство по надстройкам Word](../tutorials/word-tutorial.md)
- [Справочник по API JavaScript для Word](/javascript/api/word)
- [Сведения о программе для разработчиков Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)
