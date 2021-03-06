---
title: Объектная модель JavaScript для Word в надстройках Office
description: Сведения о важнейших классах в объектной модели JavaScript для Word.
ms.date: 10/14/2020
localization_priority: Priority
ms.openlocfilehash: 43ca88e7899e2ff11748dc91d5c8a5059d8bb559
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53077234"
---
# <a name="word-javascript-object-model-in-office-add-ins"></a>Объектная модель JavaScript для Word в надстройках Office

В этой статье описаны основные принципы использования [API JavaScript для Word](../reference/overview/word-add-ins-reference-overview.md) для создания надстроек. Представлены основные понятия, важные для использования API.

> [!IMPORTANT]
> Сведения об асинхронном типе API-интерфейсов Word и принципах их работы с документами см. в статье [Использование модели API, зависящей от приложения](../develop/application-specific-api-model.md).

## <a name="officejs-apis-for-word"></a>API-интерфейсы Office.js для Word

Надстройка Word взаимодействует с объектами в Word с помощью API JavaScript для Office, включающего две объектных модели JavaScript:

* **API JavaScript для Word**. [API-интерфейс JavaScript для Word](../reference/overview/word-add-ins-reference-overview.md) предоставляет строго типизированные объекты, с помощью которых можно получать доступ к документам, диапазонам, таблицам, спискам, форматированию и другим объектам.

* **Общие API-интерфейсы**. [Общий API](/javascript/api/office) можно использовать для доступа к таким компонентам, как пользовательский интерфейс, диалоговые окна и параметры клиентов, общие для нескольких типов приложений Office.

Скорее всего, вы будете разрабатывать большую часть функций надстроек для Word с помощью API JavaScript для Word, но вам также потребуются объекты из общего API. Например:

* [Context](/javascript/api/office/office.context). объект `Context` представляет среду выполнения надстройки и предоставляет доступ к ключевым объектам API. Он состоит из данных конфигурации документа, например `contentLanguage` и `officeTheme`, а также предоставляет сведения о среде выполнения надстройки, например `host` и `platform`. Кроме того, он предоставляет метод `requirements.isSetSupported()`, с помощью которого можно проверить, поддерживается ли указанный набор обязательных элементов приложением Excel, в котором работает надстройка.
* [Document](/javascript/api/office/office.document). Объект `Document` предоставляет метод `getFileAsync()`, позволяющий загрузить файл Word, в котором работает надстройка.

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
