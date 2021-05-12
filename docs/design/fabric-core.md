---
title: Fabric Core в Office надстройки
description: Обзор использования компонентов Fabric Core и пользовательского интерфейса Fabric в Office надстройки.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: e93efaea55841cc3bb6fa79ea1d1bbcaa76a4d05
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330204"
---
# <a name="fabric-core-in-office-add-ins"></a><span data-ttu-id="9502e-103">Fabric Core в Office надстройки</span><span class="sxs-lookup"><span data-stu-id="9502e-103">Fabric Core in Office Add-ins</span></span>

<span data-ttu-id="9502e-104">Fabric Core — это коллекция классов CSS и mixins SASS, предназначенных для использования в надстройки React *Office.* Fabric Core содержит основные элементы языка разработки пользовательского интерфейса Fluent, такие как значки, цвета, шрифты и сетки.</span><span class="sxs-lookup"><span data-stu-id="9502e-104">Fabric Core is an open-source collection of CSS classes and SASS mixins that's *intended for use in non-React* Office Add-ins. Fabric Core contains basic elements of the Fluent UI design language such as icons, colors, typefaces, and grids.</span></span> <span data-ttu-id="9502e-105">Fabric Core является независимой структурой, поэтому ее можно использовать с любым одно-страницным приложением или любой серверной веб-структурой пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="9502e-105">Fabric Core is framework independent, so it can be used with any single-page application or any server-side web UI framework.</span></span> <span data-ttu-id="9502e-106">(Он называется "Fabric Core" вместо "Fluent Core" по историческим причинам.)</span><span class="sxs-lookup"><span data-stu-id="9502e-106">(It's called "Fabric Core" instead of "Fluent Core" for historical reasons.)</span></span>

<span data-ttu-id="9502e-107">Если пользовательский интерфейс надстройки не React, вы также можете использовать набор компонентов, не React.</span><span class="sxs-lookup"><span data-stu-id="9502e-107">If your add-in's UI is not React-based, you can also make use of a set of non-React components.</span></span> <span data-ttu-id="9502e-108">См. [Office UI Fabric компоненты JS.](#use-office-ui-fabric-js-components)</span><span class="sxs-lookup"><span data-stu-id="9502e-108">See [Use Office UI Fabric JS components](#use-office-ui-fabric-js-components).</span></span>

> [!NOTE]
> <span data-ttu-id="9502e-109">В этой статье описывается использование Fabric Core в контексте Office надстройки. Но он также используется в широком диапазоне Microsoft 365 приложений и расширений.</span><span class="sxs-lookup"><span data-stu-id="9502e-109">This article describes the use of Fabric Core in the context of Office Add-ins. But it's also used in a wide range of Microsoft 365 apps and extensions.</span></span> <span data-ttu-id="9502e-110">Дополнительные сведения см. в [материале Fabric Core](https://developer.microsoft.com/fluentui#/get-started/web#fabric-core) и репо [с открытым исходным кодом Office UI Fabric Core](https://github.com/OfficeDev/office-ui-fabric-core).</span><span class="sxs-lookup"><span data-stu-id="9502e-110">For more information, see [Fabric Core](https://developer.microsoft.com/fluentui#/get-started/web#fabric-core) and the open source repo [Office UI Fabric Core](https://github.com/OfficeDev/office-ui-fabric-core).</span></span>

## <a name="use-fabric-core-icons-fonts-colors"></a><span data-ttu-id="9502e-111">Использование Fabric Core: значки, шрифты, цвета</span><span class="sxs-lookup"><span data-stu-id="9502e-111">Use Fabric Core: icons, fonts, colors</span></span>

<span data-ttu-id="9502e-112">Чтобы начать работу с Fabric Core:</span><span class="sxs-lookup"><span data-stu-id="9502e-112">To get started using Fabric Core:</span></span>

1. <span data-ttu-id="9502e-113">Добавьте ссылку CDN в HTML-код на своей странице.</span><span class="sxs-lookup"><span data-stu-id="9502e-113">Add the CDN reference to the HTML on your page.</span></span>  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css">
    ```

2. <span data-ttu-id="9502e-114">Используйте значки и шрифты Fabric Core.</span><span class="sxs-lookup"><span data-stu-id="9502e-114">Use Fabric Core icons and fonts.</span></span>

    <span data-ttu-id="9502e-115">Чтобы использовать значок Fabric Core, включите элемент "i" на странице, а затем со ссылкой на соответствующие классы.</span><span class="sxs-lookup"><span data-stu-id="9502e-115">To use a Fabric Core icon, include the "i" element on your page, and then reference the appropriate classes.</span></span> <span data-ttu-id="9502e-116">Вы можете изменять размер значка, изменяя размер шрифта.</span><span class="sxs-lookup"><span data-stu-id="9502e-116">You can control the size of the icon by changing the font size.</span></span> <span data-ttu-id="9502e-117">Например, ниже показано, как сделать очень большой значок таблицы, который использует цвет themePrimary (#0078d7).</span><span class="sxs-lookup"><span data-stu-id="9502e-117">For example, the following code shows how to make an extra-large table icon that uses the themePrimary (#0078d7) color.</span></span>

    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    <span data-ttu-id="9502e-118">Дополнительные инструкции см. в [книге Fluent UI Icons.](https://developer.microsoft.com/fluentui#/styles/web/icons)</span><span class="sxs-lookup"><span data-stu-id="9502e-118">For more detailed instructions, see [Fluent UI Icons](https://developer.microsoft.com/fluentui#/styles/web/icons).</span></span> <span data-ttu-id="9502e-119">Чтобы найти дополнительные значки, доступные в Fabric Core, используйте функцию поиска на этой странице.</span><span class="sxs-lookup"><span data-stu-id="9502e-119">To find more icons that are available in Fabric Core, use the search feature on that page.</span></span> <span data-ttu-id="9502e-120">Когда вы найдете значок для надстройки, добавьте к его имени префикс `ms-Icon--`.</span><span class="sxs-lookup"><span data-stu-id="9502e-120">When you find an icon to use in your add-in, be sure to prefix the icon name with `ms-Icon--`.</span></span>

    <span data-ttu-id="9502e-121">Сведения о размерах шрифтов и цветах, доступных в Fabric Core, см. в [typography](https://developer.microsoft.com/fluentui#/styles/web/typography) и таблице **Цветов** содержимого [в Цветах.](https://developer.microsoft.com/fluentui#/styles/web/colors)</span><span class="sxs-lookup"><span data-stu-id="9502e-121">For information about font sizes and colors that are available in Fabric Core, see [Typography](https://developer.microsoft.com/fluentui#/styles/web/typography) and the **Colors** table of contents at [Colors](https://developer.microsoft.com/fluentui#/styles/web/colors).</span></span>

<span data-ttu-id="9502e-122">Примеры, включенные в [примеры, далее](#samples) в этой статье.</span><span class="sxs-lookup"><span data-stu-id="9502e-122">Examples are included in the [Samples](#samples) later in this article.</span></span>

## <a name="use-office-ui-fabric-js-components"></a><span data-ttu-id="9502e-123">Использование Office UI Fabric JS</span><span class="sxs-lookup"><span data-stu-id="9502e-123">Use Office UI Fabric JS components</span></span>

<span data-ttu-id="9502e-124">Надстройки с React UIS также могут использовать все многие компоненты [из Office UI Fabric JS,](https://github.com/OfficeDev/office-ui-fabric-js)включая кнопки, диалоги, выборщики и многое другое.</span><span class="sxs-lookup"><span data-stu-id="9502e-124">Add-ins with non-React UIs can also use any of the many components from [Office UI Fabric JS](https://github.com/OfficeDev/office-ui-fabric-js), including buttons, dialogs, pickers, and more.</span></span> <span data-ttu-id="9502e-125">Ознакомьтесь с чтением репо для инструкций.</span><span class="sxs-lookup"><span data-stu-id="9502e-125">See the readme of the repo for instructions.</span></span>

<span data-ttu-id="9502e-126">Примеры, включенные в [примеры, далее](#samples) в этой статье.</span><span class="sxs-lookup"><span data-stu-id="9502e-126">Examples are included in the [Samples](#samples) later in this article.</span></span>

## <a name="samples"></a><span data-ttu-id="9502e-127">Примеры</span><span class="sxs-lookup"><span data-stu-id="9502e-127">Samples</span></span>

<span data-ttu-id="9502e-128">В следующем примере надстройки используют компоненты Fabric Core Office UI Fabric JS.</span><span class="sxs-lookup"><span data-stu-id="9502e-128">The following sample add-ins use Fabric Core and/or Office UI Fabric JS components.</span></span> <span data-ttu-id="9502e-129">Некоторые из этих репозитов архивироваться, что означает, что они больше не обновляются с помощью ошибок или исправлений безопасности, но вы все еще можете использовать их, чтобы узнать, как использовать компоненты пользовательского интерфейса Fabric Core и Fabric.</span><span class="sxs-lookup"><span data-stu-id="9502e-129">Some of these repos are archived, meaning that they are no longer being updated with bug or security fixes, but you can still use them to learn how to use Fabric Core and Fabric UI components.</span></span>

- [<span data-ttu-id="9502e-130">Excel Надстройка JavaScript SalesTracker</span><span class="sxs-lookup"><span data-stu-id="9502e-130">Excel Add-in JavaScript SalesTracker</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
- [<span data-ttu-id="9502e-131">Excel Надстройки SalesLeads</span><span class="sxs-lookup"><span data-stu-id="9502e-131">Excel Add-in SalesLeads</span></span>](https://github.com/OfficeDev/Excel-Add-in-SalesLeads)
- [<span data-ttu-id="9502e-132">Excel Тенденции расходов надстройки WoodGrove</span><span class="sxs-lookup"><span data-stu-id="9502e-132">Excel Add-in WoodGrove Expense Trends</span></span>](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends)
- [<span data-ttu-id="9502e-133">Excel Надстройка контента Humongous Insurance</span><span class="sxs-lookup"><span data-stu-id="9502e-133">Excel Content Add-in Humongous Insurance</span></span>](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance)
- [<span data-ttu-id="9502e-134">Office Пример пользовательского интерфейса fabric надстройки</span><span class="sxs-lookup"><span data-stu-id="9502e-134">Office Add-in Fabric UI Sample</span></span>](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample)
- [<span data-ttu-id="9502e-135">Office-Add-in-UX-Design-Patterns-Code</span><span class="sxs-lookup"><span data-stu-id="9502e-135">Office-Add-in-UX-Design-Patterns-Code</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [<span data-ttu-id="9502e-136">Outlook Надстройка GifMe</span><span class="sxs-lookup"><span data-stu-id="9502e-136">Outlook Add-in GifMe</span></span>](https://github.com/OfficeDev/Outlook-Add-in-GifMe)
- [<span data-ttu-id="9502e-137">PowerPoint Надстройка Microsoft Graph ASPNET InsertChart</span><span class="sxs-lookup"><span data-stu-id="9502e-137">PowerPoint Add-in Microsoft Graph ASPNET InsertChart</span></span>](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
- [<span data-ttu-id="9502e-138">Word Add-in Angular2 StyleChecker</span><span class="sxs-lookup"><span data-stu-id="9502e-138">Word Add-in Angular2 StyleChecker</span></span>](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)
- [<span data-ttu-id="9502e-139">Word Add-in JS Redact</span><span class="sxs-lookup"><span data-stu-id="9502e-139">Word Add-in JS Redact</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-Redact)
- [<span data-ttu-id="9502e-140">Word Add-in MarkdownConversion</span><span class="sxs-lookup"><span data-stu-id="9502e-140">Word Add-in MarkdownConversion</span></span>](https://github.com/OfficeDev/Word-Add-in-MarkdownConversion)
