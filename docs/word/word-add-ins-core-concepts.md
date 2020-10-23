---
title: Объектная модель JavaScript для Word в надстройках Office
description: Сведения о важнейших классах в объектной модели JavaScript для Word.
ms.date: 10/14/2020
localization_priority: Priority
ms.openlocfilehash: c85c56987ef5de7c087064ac668f137326089642
ms.sourcegitcommit: 42e6cfe51d99d4f3f05a3245829d764b28c46bbb
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/23/2020
ms.locfileid: "48740870"
---
# <a name="word-javascript-object-model-in-office-add-ins"></a><span data-ttu-id="05ee0-103">Объектная модель JavaScript для Word в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="05ee0-103">Word JavaScript object model in Office Add-ins</span></span>

<span data-ttu-id="05ee0-104">В этой статье описаны основные принципы использования [API JavaScript для Word](../reference/overview/word-add-ins-reference-overview.md) для создания надстроек. Представлены основные понятия, важные для использования API.</span><span class="sxs-lookup"><span data-stu-id="05ee0-104">This article describes concepts that are fundamental to using the [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) to build add-ins. It introduces core concepts that are fundamental to using the API.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="05ee0-105">Сведения об асинхронном типе API-интерфейсов Word и принципах их работы с документами см. в статье [Использование модели API, зависящей от приложения](../develop/application-specific-api-model.md).</span><span class="sxs-lookup"><span data-stu-id="05ee0-105">See [Using the application-specific API model](../develop/application-specific-api-model.md) to learn about the asynchronous nature of the Word APIs and how they work with the document.</span></span>

## <a name="officejs-apis-for-word"></a><span data-ttu-id="05ee0-106">API-интерфейсы Office.js для Word</span><span class="sxs-lookup"><span data-stu-id="05ee0-106">Office.js APIs for Word</span></span>

<span data-ttu-id="05ee0-107">Надстройка Word взаимодействует с объектами в Word с помощью API JavaScript для Office, включающего две объектных модели JavaScript:</span><span class="sxs-lookup"><span data-stu-id="05ee0-107">A Word add-in interacts with objects in Excel by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="05ee0-108">**API JavaScript для Word**. [API-интерфейс JavaScript для Word](../reference/overview/word-add-ins-reference-overview.md) предоставляет строго типизированные объекты, с помощью которых можно получать доступ к документам, диапазонам, таблицам, спискам, форматированию и другим объектам.</span><span class="sxs-lookup"><span data-stu-id="05ee0-108">**Word JavaScript API**: The [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) provides strongly-typed objects that you can use to access the document, ranges, tables, lists, formatting, and more.</span></span>

* <span data-ttu-id="05ee0-109">**Общие API-интерфейсы**. [Общий API](/javascript/api/office) можно использовать для доступа к таким компонентам, как пользовательский интерфейс, диалоговые окна и параметры клиентов, общие для нескольких типов приложений Office.</span><span class="sxs-lookup"><span data-stu-id="05ee0-109">**Common APIs**: The [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="05ee0-110">Скорее всего, вы будете разрабатывать большую часть функций надстроек для Word с помощью API JavaScript для Word, но вам также потребуются объекты из общего API.</span><span class="sxs-lookup"><span data-stu-id="05ee0-110">While you'll likely use the Word JavaScript API to develop the majority of functionality in add-ins that target Word, you'll also use objects in the Common API.</span></span> <span data-ttu-id="05ee0-111">Пример.</span><span class="sxs-lookup"><span data-stu-id="05ee0-111">For example:</span></span>

* <span data-ttu-id="05ee0-112">[Context](/javascript/api/office/office.context). объект `Context` представляет среду выполнения надстройки и предоставляет доступ к ключевым объектам API.</span><span class="sxs-lookup"><span data-stu-id="05ee0-112">[Context](/javascript/api/office/office.context): The `Context` object represents the runtime environment of the add-in and provides access to key objects of the API.</span></span> <span data-ttu-id="05ee0-113">Он состоит из данных конфигурации документа, например `contentLanguage` и `officeTheme`, а также предоставляет сведения о среде выполнения надстройки, например `host` и `platform`.</span><span class="sxs-lookup"><span data-stu-id="05ee0-113">It consists of document configuration details such as `contentLanguage` and `officeTheme` and also provides information about the add-in's runtime environment such as `host` and `platform`.</span></span> <span data-ttu-id="05ee0-114">Кроме того, он предоставляет метод `requirements.isSetSupported()`, с помощью которого можно проверить, поддерживается ли указанный набор обязательных элементов приложением Excel, в котором работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="05ee0-114">Additionally, it provides the `requirements.isSetSupported()` method, which you can use to check whether a specified requirement set is supported by the Excel application where the add-in is running.</span></span>
* <span data-ttu-id="05ee0-115">[Document](/javascript/api/office/office.document). Объект `Document` предоставляет метод `getFileAsync()`, позволяющий загрузить файл Word, в котором работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="05ee0-115">[Document](/javascript/api/office/office.document): The `Document` object provides the `getFileAsync()` method, which you can use to download the Word file where the add-in is running.</span></span>

![Изображение различий между API JS для Word и общими API](../images/word-js-api-common-api.png)

## <a name="word-specific-object-model"></a><span data-ttu-id="05ee0-117">Объектная модель для Word</span><span class="sxs-lookup"><span data-stu-id="05ee0-117">Word-specific object model</span></span>

<span data-ttu-id="05ee0-118">Чтобы понять API-интерфейсы Word, нужно понимать, как компоненты документа связаны друг с другом.</span><span class="sxs-lookup"><span data-stu-id="05ee0-118">To understand the Word APIs, you must understand how the components of a document are related to one another.</span></span>

* <span data-ttu-id="05ee0-119">Объект **Document** содержит объекты **Section**, а также объекты уровня документа, например параметры и настраиваемые части XML.</span><span class="sxs-lookup"><span data-stu-id="05ee0-119">The **Document** contains the **Section**s, and document-level entities such as settings and custom XML parts.</span></span>
* <span data-ttu-id="05ee0-120">Объект **Section** содержит объект **Body**.</span><span class="sxs-lookup"><span data-stu-id="05ee0-120">A **Section** contains a **Body**.</span></span>
* <span data-ttu-id="05ee0-121">Объект **Body** предоставляет доступ к объектам **Paragraph**, **ContentControl** и **Range**, а также к другим объектам.</span><span class="sxs-lookup"><span data-stu-id="05ee0-121">A **Body** gives access to **Paragraph**s, **ContentControl**s, and **Range** objects, among others.</span></span>
* <span data-ttu-id="05ee0-122">Объект **Range** представляет собой непрерывную область содержимого, включающую текст, пробелы, объекты **Table**, а также изображения.</span><span class="sxs-lookup"><span data-stu-id="05ee0-122">A **Range** represents a contiguous area of content, including text, white space, **Table**s, and images.</span></span> <span data-ttu-id="05ee0-123">Он также содержит большую часть методов обработки текста.</span><span class="sxs-lookup"><span data-stu-id="05ee0-123">It also contains most of the text manipulation methods.</span></span>
* <span data-ttu-id="05ee0-124">Объект **List** представляет текст в виде нумерованного или маркированного списка.</span><span class="sxs-lookup"><span data-stu-id="05ee0-124">A **List** represents text in a numbered or bulleted list.</span></span>

## <a name="see-also"></a><span data-ttu-id="05ee0-125">См. также</span><span class="sxs-lookup"><span data-stu-id="05ee0-125">See also</span></span>

- [<span data-ttu-id="05ee0-126">Обзор API JavaScript для Word</span><span class="sxs-lookup"><span data-stu-id="05ee0-126">Word JavaScript API overview</span></span>](../reference/overview/word-add-ins-reference-overview.md)
- [<span data-ttu-id="05ee0-127">Создание первой надстройки Word</span><span class="sxs-lookup"><span data-stu-id="05ee0-127">Build your first Word add-in</span></span>](../quickstarts/word-quickstart.md)
- [<span data-ttu-id="05ee0-128">Руководство по надстройкам Word</span><span class="sxs-lookup"><span data-stu-id="05ee0-128">Word add-in tutorial</span></span>](../tutorials/word-tutorial.md)
- [<span data-ttu-id="05ee0-129">Справочник по API JavaScript для Word</span><span class="sxs-lookup"><span data-stu-id="05ee0-129">Word JavaScript API reference</span></span>](/javascript/api/word)
- [<span data-ttu-id="05ee0-130">Сведения о программе для разработчиков Microsoft 365</span><span class="sxs-lookup"><span data-stu-id="05ee0-130">Learn about the Microsoft 365 Developer Program</span></span>](https://developer.microsoft.com/microsoft-365/dev-program)