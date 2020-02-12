---
title: Office UI Fabric в надстройках Office 
description: ''
ms.date: 12/04/2017
localization_priority: Normal
ms.openlocfilehash: ddec544b1aac146e75e8b22eba95e243b391d70f
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950476"
---
# <a name="office-ui-fabric-in-office-add-ins"></a><span data-ttu-id="1c5bf-102">Office UI Fabric в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="1c5bf-102">Office UI Fabric in Office Add-ins</span></span> 

<span data-ttu-id="1c5bf-p101">Office UI Fabric — это интерфейсная платформа JavaScript для создания дизайна для Office и Office 365. В Fabric предоставлены компоненты дизайна, которые можно расширять, дорабатывать и использовать в надстройке Office. Так как Fabric использует язык дизайна Office, компоненты дизайна Fabric выглядят в Office очень естественно.</span><span class="sxs-lookup"><span data-stu-id="1c5bf-p101">Office UI Fabric is a JavaScript front-end framework for building user experiences for Office and Office 365. Fabric provides visuals-focused components that you can extend, rework, and use in your Office Add-in. Because Fabric uses the Office Design Language, Fabric's UX components look like a natural extension of Office.</span></span> 

<span data-ttu-id="1c5bf-p102">Рекомендуем использовать Office UI Fabric для создания надстроек. Использовать Office UI Fabric необязательно.</span><span class="sxs-lookup"><span data-stu-id="1c5bf-p102">If you are building an add-in, we encourage you to use Office UI Fabric to create your user experience. Using Office UI Fabric is optional.</span></span>

<span data-ttu-id="1c5bf-108">В следующих разделах описано, как начать использовать Fabric для своих потребностей.</span><span class="sxs-lookup"><span data-stu-id="1c5bf-108">The following sections explain how to get started using Fabric to meet your requirements.</span></span> 

## <a name="use-fabric-core-icons-fonts-colors"></a><span data-ttu-id="1c5bf-109">Использование Fabric Core: значки, шрифты, цвета</span><span class="sxs-lookup"><span data-stu-id="1c5bf-109">Use Fabric Core: icons, fonts, colors</span></span>
<span data-ttu-id="1c5bf-110">Fabric Core содержит основные элементы языка дизайна, такие как значки, цвета, тип и сетку.</span><span class="sxs-lookup"><span data-stu-id="1c5bf-110">Fabric Core contains basic elements of the design language such as icons, colors, type, and grid.</span></span><span data-ttu-id="1c5bf-111">Fabric Core не зависит от платформы.</span><span class="sxs-lookup"><span data-stu-id="1c5bf-111"> Fabric core is framework independent.</span></span> <span data-ttu-id="1c5bf-112">Fabric Core используется с помощью Fabric React и входит в его состав.</span><span class="sxs-lookup"><span data-stu-id="1c5bf-112">Fabric Core is used by and included with Fabric React.</span></span>

<span data-ttu-id="1c5bf-113">Чтобы начать работу с Fabric Core:</span><span class="sxs-lookup"><span data-stu-id="1c5bf-113">To get started using Fabric Core:</span></span>

1. <span data-ttu-id="1c5bf-114">Добавьте ссылку CDN в HTML-код на своей странице.</span><span class="sxs-lookup"><span data-stu-id="1c5bf-114">Add the CDN reference to the HTML on your page.</span></span>  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css">
    ```   
    
2. <span data-ttu-id="1c5bf-115">Используйте значки и шрифты Fabric.</span><span class="sxs-lookup"><span data-stu-id="1c5bf-115">Use Fabric icons and fonts.</span></span> 

    <span data-ttu-id="1c5bf-p104">Чтобы использовать значок Fabric, разместите элемент "i" на своей странице и сошлитесь на соответствующие классы. Вы можете сами выбирать размер значка, изменяя размер шрифта. Например, в коде ниже показано, как сделать очень большой значок таблицы, который использует цвет themePrimary (#0078d7).</span><span class="sxs-lookup"><span data-stu-id="1c5bf-p104">To use a Fabric icon, include the "i" element on your page, and then reference the appropriate classes. You can control the size of the icon by changing the font size. For example, the following code shows how to make an extra-large table icon that uses the themePrimary (#0078d7) color.</span></span> 
   
    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    <span data-ttu-id="1c5bf-p105">Чтобы найти другие значки, доступные в Office UI Fabric, используйте функцию поиска на странице [Значки](https://developer.microsoft.com/fabric#/styles/icons). Когда вы найдете значок для надстройки, добавьте к его имени префикс `ms-Icon--`.</span><span class="sxs-lookup"><span data-stu-id="1c5bf-p105">To find more icons that are available in Office UI Fabric, use the search feature on the [Icons](https://developer.microsoft.com/fabric#/styles/icons) page. When you find an icon to use in your add-in, be sure to prefix the icon name with `ms-Icon--`.</span></span> 

    <span data-ttu-id="1c5bf-121">Сведения о размерах шрифтов и цветах, доступных в Office UI Fabric, см. в разделах [Оформление](https://developer.microsoft.com/fabric#/styles/typography) и [Цвета](https://developer.microsoft.com/fabric#/styles/colors).</span><span class="sxs-lookup"><span data-stu-id="1c5bf-121">For information about font sizes and colors that are available in Office UI Fabric, see [Typography](https://developer.microsoft.com/fabric#/styles/typography) and [Colors](https://developer.microsoft.com/fabric#/styles/colors).</span></span>
 
## <a name="use-fabric-components"></a><span data-ttu-id="1c5bf-122">Использование компонентов Fabric</span><span class="sxs-lookup"><span data-stu-id="1c5bf-122">Use Fabric Components</span></span> 
<span data-ttu-id="1c5bf-123">В Fabric есть различные компоненты оформления, которые можно использовать при создании надстроек, в том числе:</span><span class="sxs-lookup"><span data-stu-id="1c5bf-123">Fabric provides a variety of UX components that you can use to build your add-in, including the following types of components:</span></span>

- <span data-ttu-id="1c5bf-124">Компоненты ввода — например, Button, Checkbox и Toggle.</span><span class="sxs-lookup"><span data-stu-id="1c5bf-124">Input components - for example, Button, Checkbox, and Toggle</span></span>
- <span data-ttu-id="1c5bf-125">Компоненты навигации — например, Pivot и Breadcrumb.</span><span class="sxs-lookup"><span data-stu-id="1c5bf-125">Navigation components - for example, Pivot and Breadcrumb</span></span>
- <span data-ttu-id="1c5bf-126">Компоненты уведомления — например, MessageBar и Callout.</span><span class="sxs-lookup"><span data-stu-id="1c5bf-126">Notification components - for example, MessageBar and Callout</span></span>  

<span data-ttu-id="1c5bf-127">Не все компоненты Fabric рекомендуется использовать в надстройках. Ниже приведен список компонентов дизайна Fabric React, рекомендованных для надстроек.</span><span class="sxs-lookup"><span data-stu-id="1c5bf-127">Not all Fabric components are recommended for use in add-ins. Here is a list of Fabric React UX components that we recommend for use in an add-in:</span></span>

- [<span data-ttu-id="1c5bf-128">Строка навигации</span><span class="sxs-lookup"><span data-stu-id="1c5bf-128">Breadcrumb</span></span>](https://developer.microsoft.com/fabric#/components/breadcrumb)
- [<span data-ttu-id="1c5bf-129">Кнопка</span><span class="sxs-lookup"><span data-stu-id="1c5bf-129">Button</span></span>](https://developer.microsoft.com/fabric#/components/button)
- [<span data-ttu-id="1c5bf-130">Флажок</span><span class="sxs-lookup"><span data-stu-id="1c5bf-130">Checkbox</span></span>](https://developer.microsoft.com/fabric#/components/checkbox)
- [<span data-ttu-id="1c5bf-131">ChoiceGroup</span><span class="sxs-lookup"><span data-stu-id="1c5bf-131">ChoiceGroup</span></span>](https://developer.microsoft.com/fabric#/components/choicegroup)
- [<span data-ttu-id="1c5bf-132">Раскрывающееся меню</span><span class="sxs-lookup"><span data-stu-id="1c5bf-132">Dropdown</span></span>](https://developer.microsoft.com/fabric#/components/dropdown)
- [<span data-ttu-id="1c5bf-133">Подпись</span><span class="sxs-lookup"><span data-stu-id="1c5bf-133">Label</span></span>](https://developer.microsoft.com/fabric#/components/label)
- [<span data-ttu-id="1c5bf-134">Список</span><span class="sxs-lookup"><span data-stu-id="1c5bf-134">List</span></span>](https://developer.microsoft.com/fabric#/components/list)
- [<span data-ttu-id="1c5bf-135">Сводка</span><span class="sxs-lookup"><span data-stu-id="1c5bf-135">Pivot</span></span>](https://developer.microsoft.com/fabric#/components/pivot)
- [<span data-ttu-id="1c5bf-136">TextField</span><span class="sxs-lookup"><span data-stu-id="1c5bf-136">TextField</span></span>](https://developer.microsoft.com/fabric#/components/textfield)
- [<span data-ttu-id="1c5bf-137">Переключатель</span><span class="sxs-lookup"><span data-stu-id="1c5bf-137">Toggle</span></span>](https://developer.microsoft.com/fabric#/components/toggle)

<span data-ttu-id="1c5bf-p106">Для создания надстройки можно использовать разные платформы JavaScript, такие как Angular или React. Прежде чем использовать компоненты Fabric со своей платформой, ознакомьтесь с перечисленными ниже ресурсами.</span><span class="sxs-lookup"><span data-stu-id="1c5bf-p106">You can use different JavaScript frameworks, such as Angular or React, to build your add-in. To get started using Fabric components with your framework, see the following resources.</span></span>

|<span data-ttu-id="1c5bf-140">**Платформа**</span><span class="sxs-lookup"><span data-stu-id="1c5bf-140">**Framework**</span></span>|<span data-ttu-id="1c5bf-141">**Пример**</span><span class="sxs-lookup"><span data-stu-id="1c5bf-141">**Example**</span></span>|
|:------------|:----------|
|<span data-ttu-id="1c5bf-142">**React**</span><span class="sxs-lookup"><span data-stu-id="1c5bf-142">**React**</span></span>|[<span data-ttu-id="1c5bf-143">Использование Office UI Fabric React в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="1c5bf-143">Using Office UI Fabric React in Office Add-ins</span></span>](using-office-ui-fabric-react.md )|
|<span data-ttu-id="1c5bf-144">**Angular**</span><span class="sxs-lookup"><span data-stu-id="1c5bf-144">**Angular**</span></span>| <span data-ttu-id="1c5bf-145">См. проект сообщества [ngOfficeUIFabric](http://ngofficeuifabric.com/) с директивами Angular 1.5 и раздел [Обдумайте размещение компонентов Fabric в компонентах Angular 2](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components)</span><span class="sxs-lookup"><span data-stu-id="1c5bf-145">See [ngOfficeUIFabric](http://ngofficeuifabric.com/), which is a community project with Angular 1.5 directives, and [Consider wrapping Fabric components with Angular 2 components](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components)</span></span>|
