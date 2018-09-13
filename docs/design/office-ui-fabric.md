---
title: Office UI Fabric в надстройках Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 7b1e4a9c377c9a60195a51115d7f275603f1ca5a
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2018
ms.locfileid: "23944036"
---
# <a name="office-ui-fabric-in-office-add-ins"></a><span data-ttu-id="d0dee-102">Office UI Fabric в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="d0dee-102">Office UI Fabric in Office Add-ins</span></span> 

<span data-ttu-id="d0dee-p101">Office UI Fabric — это интерфейсная платформа JavaScript для создания дизайна для Office и Office 365. В Fabric предоставлены компоненты дизайна, которые можно расширять, дорабатывать и использовать в надстройке Office. Так как Fabric использует язык дизайна Office, компоненты дизайна Fabric выглядят в Office очень естественно.</span><span class="sxs-lookup"><span data-stu-id="d0dee-p101">Office UI Fabric is a JavaScript front-end framework for building user experiences for Office and Office 365. Fabric provides visuals-focused components that you can extend, rework, and use in your Office Add-in. Because Fabric uses the Office Design Language, Fabric's UX components look like a natural extension of Office.</span></span> 

<span data-ttu-id="d0dee-p102">Рекомендуем использовать Office UI Fabric для создания надстроек. Использовать Office UI Fabric необязательно.</span><span class="sxs-lookup"><span data-stu-id="d0dee-p102">If you are building an add-in, we encourage you to use Office UI Fabric to create your user experience. Using Office UI Fabric is optional.</span></span>

<span data-ttu-id="d0dee-108">В следующих разделах описано, как начать использовать Fabric для своих потребностей.</span><span class="sxs-lookup"><span data-stu-id="d0dee-108">The following sections explain how to get started using Fabric to meet your requirements.</span></span> 

## <a name="use-fabric-core-icons-fonts-colors"></a><span data-ttu-id="d0dee-109">Использование Fabric Core: значки, шрифты, цвета</span><span class="sxs-lookup"><span data-stu-id="d0dee-109">Use Fabric Core: icons, fonts, colors</span></span>
<span data-ttu-id="d0dee-p103">Fabric Core содержит основные элементы языка дизайна, такие как значки, цвета, тип и сетку. Fabric Core не зависит от платформы. И Fabric React, и Fabric JS используют Fabric Core.</span><span class="sxs-lookup"><span data-stu-id="d0dee-p103">Fabric Core contains basic elements of the design language such as icons, colors, type, and grid. Fabric core is framework independent. Both Fabric React and Fabric JS use Fabric Core.</span></span>

<span data-ttu-id="d0dee-113">Чтобы начать работу с Fabric Core:</span><span class="sxs-lookup"><span data-stu-id="d0dee-113">To get started using Fabric Core:</span></span>

1. <span data-ttu-id="d0dee-114">Добавьте ссылку CDN в HTML-код на своей странице.</span><span class="sxs-lookup"><span data-stu-id="d0dee-114">Add the CDN reference to the HTML on your page.</span></span>  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css">
    ```   
    
2. <span data-ttu-id="d0dee-115">Используйте значки и шрифты Fabric.</span><span class="sxs-lookup"><span data-stu-id="d0dee-115">Use Fabric icons and fonts.</span></span> 

    <span data-ttu-id="d0dee-p104">Чтобы использовать значок Fabric, разместите элемент "i" на своей странице и сошлитесь на соответствующие классы. Вы можете сами выбирать размер значка, изменяя размер шрифта. Например, в коде ниже показано, как сделать очень большой значок таблицы, который использует цвет themePrimary (#0078d7).</span><span class="sxs-lookup"><span data-stu-id="d0dee-p104">To use a Fabric icon, include the "i" element on your page, and then reference the appropriate classes. You can control the size of the icon by changing the font size. For example, the following code shows how to make an extra-large table icon that uses the themePrimary (#0078d7) color.</span></span> 
   
    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    <span data-ttu-id="d0dee-p105">Чтобы найти другие значки, доступные в Office UI Fabric, используйте функцию поиска на странице [Значки](https://developer.microsoft.com/fabric#/styles/icons). Когда вы найдете значок для надстройки, добавьте к его имени префикс `ms-Icon--`.</span><span class="sxs-lookup"><span data-stu-id="d0dee-p105">To find more icons that are available in Office UI Fabric, use the search feature on the [Icons](https://developer.microsoft.com/fabric#/styles/icons) page. When you find an icon to use in your add-in, be sure to prefix the icon name with `ms-Icon--`.</span></span> 

    <span data-ttu-id="d0dee-121">Сведения о размерах шрифтов и цветах, доступных в Office UI Fabric, см. в разделах [Оформление](https://developer.microsoft.com/fabric#/styles/typography) и [Цвета](https://developer.microsoft.com/fabric#/styles/colors).</span><span class="sxs-lookup"><span data-stu-id="d0dee-121">For information about font sizes and colors that are available in Office UI Fabric, see [Typography](https://developer.microsoft.com/fabric#/styles/typography) and [Colors](https://developer.microsoft.com/fabric#/styles/colors).</span></span>
 
## <a name="use-fabric-components"></a><span data-ttu-id="d0dee-122">Использование компонентов Fabric</span><span class="sxs-lookup"><span data-stu-id="d0dee-122">Use Fabric Components</span></span> 
<span data-ttu-id="d0dee-123">В Fabric есть различные компоненты оформления, которые можно использовать при создании надстроек, в том числе:</span><span class="sxs-lookup"><span data-stu-id="d0dee-123">Fabric provides a variety of UX components that you can use to build your add-in, including the following types of components:</span></span>

- <span data-ttu-id="d0dee-124">Компоненты ввода — например, Button, Checkbox и Toggle.</span><span class="sxs-lookup"><span data-stu-id="d0dee-124">Input components - for example, Button, Checkbox, and Toggle</span></span>
- <span data-ttu-id="d0dee-125">Компоненты навигации – например, сводный документ и строка навигации</span><span class="sxs-lookup"><span data-stu-id="d0dee-125">Navigation components - for example, Pivot Breadcrumb</span></span>
- <span data-ttu-id="d0dee-126">Компоненты уведомления — например, MessageBar и Callout.</span><span class="sxs-lookup"><span data-stu-id="d0dee-126">Notification components - for example, MessageBar and Callout</span></span>  

<span data-ttu-id="d0dee-127">Не все компоненты Fabric рекомендуются для использования в надстройках. Ниже приведен список компонентов Fabric React UX, которые мы рекомендуем использовать в надстройке:</span><span class="sxs-lookup"><span data-stu-id="d0dee-127">Not all Fabric components are recommended for use in add-ins. Here is a list of Fabric React UX components that we recommend for use in an add-in:</span></span>

- [<span data-ttu-id="d0dee-128">Строка навигации</span><span class="sxs-lookup"><span data-stu-id="d0dee-128">Breadcrumb</span></span>](https://developer.microsoft.com/fabric#/components/breadcrumb)
- [<span data-ttu-id="d0dee-129">Кнопка</span><span class="sxs-lookup"><span data-stu-id="d0dee-129">Button</span></span>](https://developer.microsoft.com/fabric#/components/button)
- [<span data-ttu-id="d0dee-130">Флажок</span><span class="sxs-lookup"><span data-stu-id="d0dee-130">Checkbox</span></span>](https://developer.microsoft.com/fabric#/components/checkbox)
- [<span data-ttu-id="d0dee-131">ChoiceGroup</span><span class="sxs-lookup"><span data-stu-id="d0dee-131">ChoiceGroup</span></span>](https://developer.microsoft.com/fabric#/components/choicegroup)
- [<span data-ttu-id="d0dee-132">Раскрывающееся меню</span><span class="sxs-lookup"><span data-stu-id="d0dee-132">Dropdown</span></span>](https://developer.microsoft.com/fabric#/components/dropdown)
- [<span data-ttu-id="d0dee-133">Подпись</span><span class="sxs-lookup"><span data-stu-id="d0dee-133">Label</span></span>](https://developer.microsoft.com/fabric#/components/label)
- [<span data-ttu-id="d0dee-134">Список</span><span class="sxs-lookup"><span data-stu-id="d0dee-134">List</span></span>](https://developer.microsoft.com/fabric#/components/list)
- [<span data-ttu-id="d0dee-135">Сводка</span><span class="sxs-lookup"><span data-stu-id="d0dee-135">Pivot</span></span>](https://developer.microsoft.com/fabric#/components/pivot)
- [<span data-ttu-id="d0dee-136">TextField</span><span class="sxs-lookup"><span data-stu-id="d0dee-136">TextField</span></span>](https://developer.microsoft.com/fabric#/components/textfield)
- [<span data-ttu-id="d0dee-137">Переключатель</span><span class="sxs-lookup"><span data-stu-id="d0dee-137">Toggle</span></span>](https://developer.microsoft.com/fabric#/components/toggle)

<span data-ttu-id="d0dee-p106">Для создания надстройки можно использовать разные платформы JavaScript, такие как Angular или React. Прежде чем использовать компоненты Fabric со своей платформой, ознакомьтесь с перечисленными ниже ресурсами.</span><span class="sxs-lookup"><span data-stu-id="d0dee-p106">You can use different JavaScript frameworks, such as Angular or React, to build your add-in. To get started using Fabric components with your framework, see the following resources.</span></span>

|<span data-ttu-id="d0dee-140">**Платформа**</span><span class="sxs-lookup"><span data-stu-id="d0dee-140">**Framework**</span></span>|<span data-ttu-id="d0dee-141">**Пример**</span><span class="sxs-lookup"><span data-stu-id="d0dee-141">**Example**</span></span>|
|:------------|:----------|
|<span data-ttu-id="d0dee-142">**|||UNTRANSLATED_CONTENT_START|||React|||UNTRANSLATED_CONTENT_END|||**</span><span class="sxs-lookup"><span data-stu-id="d0dee-142">**React**</span></span>|[<span data-ttu-id="d0dee-143">Использование Office UI Fabric React в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="d0dee-143">Using Office UI Fabric React in Office Add-ins</span></span>](using-office-ui-fabric-react.md )|
|<span data-ttu-id="d0dee-144">**Angular**</span><span class="sxs-lookup"><span data-stu-id="d0dee-144">**Angular**</span></span>| <span data-ttu-id="d0dee-145">См. проект сообщества [ngOfficeUIFabric](http://ngofficeuifabric.com/) с директивами Angular 1.5 и раздел [Обдумайте размещение компонентов Fabric в компонентах Angular 2](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components)</span><span class="sxs-lookup"><span data-stu-id="d0dee-145">See [ngOfficeUIFabric](http://ngofficeuifabric.com/), which is a community project with Angular 1.5 directives, and [Consider wrapping Fabric components with Angular 2 components](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components)</span></span>|
