---
title: Office UI Fabric в надстройках Office 
description: Общие сведения о том, как использовать компоненты Office UI Fabric в надстройках Office.
ms.date: 04/20/2020
localization_priority: Normal
ms.openlocfilehash: 7b74c734fb2559e4bf4408e40eb5366f9b79b755
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608504"
---
# <a name="office-ui-fabric-in-office-add-ins"></a><span data-ttu-id="ad39f-103">Office UI Fabric в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="ad39f-103">Office UI Fabric in Office Add-ins</span></span> 

<span data-ttu-id="ad39f-p101">Office UI Fabric — это интерфейсная платформа JavaScript для создания дизайна для Office и Office 365. В Fabric предоставлены компоненты дизайна, которые можно расширять, дорабатывать и использовать в надстройке Office. Так как Fabric использует язык дизайна Office, компоненты дизайна Fabric выглядят в Office очень естественно.</span><span class="sxs-lookup"><span data-stu-id="ad39f-p101">Office UI Fabric is a JavaScript front-end framework for building user experiences for Office and Office 365. Fabric provides visuals-focused components that you can extend, rework, and use in your Office Add-in. Because Fabric uses the Office Design Language, Fabric's UX components look like a natural extension of Office.</span></span> 

<span data-ttu-id="ad39f-p102">Рекомендуем использовать Office UI Fabric для создания надстроек. Использовать Office UI Fabric необязательно.</span><span class="sxs-lookup"><span data-stu-id="ad39f-p102">If you are building an add-in, we encourage you to use Office UI Fabric to create your user experience. Using Office UI Fabric is optional.</span></span>

<span data-ttu-id="ad39f-109">В следующих разделах описано, как начать использовать Fabric для своих потребностей.</span><span class="sxs-lookup"><span data-stu-id="ad39f-109">The following sections explain how to get started using Fabric to meet your requirements.</span></span> 

## <a name="use-fabric-core-icons-fonts-colors"></a><span data-ttu-id="ad39f-110">Использование Fabric Core: значки, шрифты, цвета</span><span class="sxs-lookup"><span data-stu-id="ad39f-110">Use Fabric Core: icons, fonts, colors</span></span>
<span data-ttu-id="ad39f-111">Fabric Core содержит основные элементы языка дизайна, такие как значки, цвета, тип и сетку.</span><span class="sxs-lookup"><span data-stu-id="ad39f-111">Fabric Core contains basic elements of the design language such as icons, colors, type, and grid.</span></span><span data-ttu-id="ad39f-112">Fabric Core не зависит от платформы.</span><span class="sxs-lookup"><span data-stu-id="ad39f-112"> Fabric core is framework independent.</span></span> <span data-ttu-id="ad39f-113">Fabric Core используется с помощью Fabric React и входит в его состав.</span><span class="sxs-lookup"><span data-stu-id="ad39f-113">Fabric Core is used by and included with Fabric React.</span></span>

<span data-ttu-id="ad39f-114">Чтобы начать работу с Fabric Core:</span><span class="sxs-lookup"><span data-stu-id="ad39f-114">To get started using Fabric Core:</span></span>

1. <span data-ttu-id="ad39f-115">Добавьте ссылку CDN в HTML-код на своей странице.</span><span class="sxs-lookup"><span data-stu-id="ad39f-115">Add the CDN reference to the HTML on your page.</span></span>  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css">
    ```   
    
2. <span data-ttu-id="ad39f-116">Используйте значки и шрифты Fabric.</span><span class="sxs-lookup"><span data-stu-id="ad39f-116">Use Fabric icons and fonts.</span></span> 

    <span data-ttu-id="ad39f-p104">Чтобы использовать значок Fabric, разместите элемент "i" на своей странице и сошлитесь на соответствующие классы. Вы можете сами выбирать размер значка, изменяя размер шрифта. Например, в коде ниже показано, как сделать очень большой значок таблицы, который использует цвет themePrimary (#0078d7).</span><span class="sxs-lookup"><span data-stu-id="ad39f-p104">To use a Fabric icon, include the "i" element on your page, and then reference the appropriate classes. You can control the size of the icon by changing the font size. For example, the following code shows how to make an extra-large table icon that uses the themePrimary (#0078d7) color.</span></span> 
   
    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    <span data-ttu-id="ad39f-p105">Чтобы найти другие значки, доступные в Office UI Fabric, используйте функцию поиска на странице [Значки](https://developer.microsoft.com/fabric#/styles/icons). Когда вы найдете значок для надстройки, добавьте к его имени префикс `ms-Icon--`.</span><span class="sxs-lookup"><span data-stu-id="ad39f-p105">To find more icons that are available in Office UI Fabric, use the search feature on the [Icons](https://developer.microsoft.com/fabric#/styles/icons) page. When you find an icon to use in your add-in, be sure to prefix the icon name with `ms-Icon--`.</span></span> 

    <span data-ttu-id="ad39f-122">Сведения о размерах шрифтов и цветах, доступных в Office UI Fabric, см. в разделах [Оформление](https://developer.microsoft.com/fabric#/styles/typography) и [Цвета](https://developer.microsoft.com/fabric#/styles/colors).</span><span class="sxs-lookup"><span data-stu-id="ad39f-122">For information about font sizes and colors that are available in Office UI Fabric, see [Typography](https://developer.microsoft.com/fabric#/styles/typography) and [Colors](https://developer.microsoft.com/fabric#/styles/colors).</span></span>
 
## <a name="use-fabric-components"></a><span data-ttu-id="ad39f-123">Использование компонентов Fabric</span><span class="sxs-lookup"><span data-stu-id="ad39f-123">Use Fabric Components</span></span> 
<span data-ttu-id="ad39f-124">В Fabric есть различные компоненты оформления, которые можно использовать при создании надстроек, в том числе:</span><span class="sxs-lookup"><span data-stu-id="ad39f-124">Fabric provides a variety of UX components that you can use to build your add-in, including the following types of components:</span></span>

- <span data-ttu-id="ad39f-125">Компоненты ввода — например, Button, Checkbox и Toggle.</span><span class="sxs-lookup"><span data-stu-id="ad39f-125">Input components - for example, Button, Checkbox, and Toggle</span></span>
- <span data-ttu-id="ad39f-126">Компоненты навигации — например, Pivot и Breadcrumb.</span><span class="sxs-lookup"><span data-stu-id="ad39f-126">Navigation components - for example, Pivot and Breadcrumb</span></span>
- <span data-ttu-id="ad39f-127">Компоненты уведомления — например, MessageBar и Callout.</span><span class="sxs-lookup"><span data-stu-id="ad39f-127">Notification components - for example, MessageBar and Callout</span></span>  

<span data-ttu-id="ad39f-128">Не все компоненты Fabric рекомендуется использовать в надстройках. Ниже приведен список компонентов дизайна Fabric React, рекомендованных для надстроек.</span><span class="sxs-lookup"><span data-stu-id="ad39f-128">Not all Fabric components are recommended for use in add-ins. Here is a list of Fabric React UX components that we recommend for use in an add-in:</span></span>

- [<span data-ttu-id="ad39f-129">Строка навигации</span><span class="sxs-lookup"><span data-stu-id="ad39f-129">Breadcrumb</span></span>](https://developer.microsoft.com/fabric#/components/breadcrumb)
- [<span data-ttu-id="ad39f-130">Кнопка</span><span class="sxs-lookup"><span data-stu-id="ad39f-130">Button</span></span>](https://developer.microsoft.com/fabric#/components/button)
- [<span data-ttu-id="ad39f-131">Флажок</span><span class="sxs-lookup"><span data-stu-id="ad39f-131">Checkbox</span></span>](https://developer.microsoft.com/fabric#/components/checkbox)
- [<span data-ttu-id="ad39f-132">ChoiceGroup</span><span class="sxs-lookup"><span data-stu-id="ad39f-132">ChoiceGroup</span></span>](https://developer.microsoft.com/fabric#/components/choicegroup)
- [<span data-ttu-id="ad39f-133">Раскрывающееся меню</span><span class="sxs-lookup"><span data-stu-id="ad39f-133">Dropdown</span></span>](https://developer.microsoft.com/fabric#/components/dropdown)
- [<span data-ttu-id="ad39f-134">Подпись</span><span class="sxs-lookup"><span data-stu-id="ad39f-134">Label</span></span>](https://developer.microsoft.com/fabric#/components/label)
- [<span data-ttu-id="ad39f-135">Список</span><span class="sxs-lookup"><span data-stu-id="ad39f-135">List</span></span>](https://developer.microsoft.com/fabric#/components/list)
- [<span data-ttu-id="ad39f-136">Сводка</span><span class="sxs-lookup"><span data-stu-id="ad39f-136">Pivot</span></span>](https://developer.microsoft.com/fabric#/components/pivot)
- [<span data-ttu-id="ad39f-137">TextField</span><span class="sxs-lookup"><span data-stu-id="ad39f-137">TextField</span></span>](https://developer.microsoft.com/fabric#/components/textfield)
- [<span data-ttu-id="ad39f-138">Переключатель</span><span class="sxs-lookup"><span data-stu-id="ad39f-138">Toggle</span></span>](https://developer.microsoft.com/fabric#/components/toggle)

<span data-ttu-id="ad39f-p106">Для создания надстройки можно использовать разные платформы JavaScript, такие как Angular или React. Прежде чем использовать компоненты Fabric со своей платформой, ознакомьтесь с перечисленными ниже ресурсами.</span><span class="sxs-lookup"><span data-stu-id="ad39f-p106">You can use different JavaScript frameworks, such as Angular or React, to build your add-in. To get started using Fabric components with your framework, see the following resources.</span></span>

|<span data-ttu-id="ad39f-141">**Платформа**</span><span class="sxs-lookup"><span data-stu-id="ad39f-141">**Framework**</span></span>|<span data-ttu-id="ad39f-142">**Пример**</span><span class="sxs-lookup"><span data-stu-id="ad39f-142">**Example**</span></span>|
|:------------|:----------|
|<span data-ttu-id="ad39f-143">**React**</span><span class="sxs-lookup"><span data-stu-id="ad39f-143">**React**</span></span>|[<span data-ttu-id="ad39f-144">Использование Office UI Fabric React в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="ad39f-144">Using Office UI Fabric React in Office Add-ins</span></span>](using-office-ui-fabric-react.md )|
|<span data-ttu-id="ad39f-145">**Angular**</span><span class="sxs-lookup"><span data-stu-id="ad39f-145">**Angular**</span></span>| [<span data-ttu-id="ad39f-146">Рассмотрите оболочку компонентов Fabric с помощью компонентов радиального 2</span><span class="sxs-lookup"><span data-stu-id="ad39f-146">Consider wrapping Fabric components with Angular 2 components</span></span>](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components)|
