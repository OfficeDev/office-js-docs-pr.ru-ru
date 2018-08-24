---
title: Office UI Fabric в надстройках Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: b573f720ebe4f90f7d4dbfdb05693871b93a2258
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925194"
---
# <a name="office-ui-fabric-in-office-add-ins"></a><span data-ttu-id="55b5b-102">Office UI Fabric в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="55b5b-102">Office UI Fabric in Office Add-ins</span></span> 

<span data-ttu-id="55b5b-p101">Office UI Fabric — это интерфейсная платформа JavaScript для создания дизайна для Office и Office 365. В Fabric предоставлены компоненты дизайна, которые можно расширять, дорабатывать и использовать в надстройке Office. Так как Fabric использует язык дизайна Office, компоненты дизайна Fabric выглядят в Office очень естественно.</span><span class="sxs-lookup"><span data-stu-id="55b5b-p101">Office UI Fabric is a JavaScript front-end framework for building user experiences for Office and Office 365. Fabric provides visuals-focused components that you can extend, rework, and use in your Office Add-in. Because Fabric uses the Office Design Language, Fabric's UX components look like a natural extension of Office.</span></span> 

<span data-ttu-id="55b5b-p102">Рекомендуем использовать Office UI Fabric для создания надстроек. Использовать Office UI Fabric необязательно.</span><span class="sxs-lookup"><span data-stu-id="55b5b-p102">If you are building an add-in, we encourage you to use Office UI Fabric to create your user experience. Using Office UI Fabric is optional.</span></span>

<span data-ttu-id="55b5b-108">В следующих разделах описано, как начать использовать Fabric для своих потребностей.</span><span class="sxs-lookup"><span data-stu-id="55b5b-108">The following sections explain how to get started using Fabric to meet your requirements.</span></span> 

## <a name="use-fabric-core-icons-fonts-colors"></a><span data-ttu-id="55b5b-109">Использование Fabric Core: значки, шрифты, цвета</span><span class="sxs-lookup"><span data-stu-id="55b5b-109">Use Fabric Core: icons, fonts, colors</span></span>
<span data-ttu-id="55b5b-p103">Fabric Core содержит основные элементы языка дизайна, такие как значки, цвета, тип и сетку. Fabric Core не зависит от платформы. И Fabric React, и Fabric JS используют Fabric Core.</span><span class="sxs-lookup"><span data-stu-id="55b5b-p103">Fabric Core contains basic elements of the design language such as icons, colors, type, and grid. Fabric core is framework independent. Both Fabric React and Fabric JS use Fabric Core.</span></span>

<span data-ttu-id="55b5b-113">Чтобы начать работу с Fabric Core:</span><span class="sxs-lookup"><span data-stu-id="55b5b-113">To get started using Fabric Core:</span></span>

1. <span data-ttu-id="55b5b-114">Добавьте ссылку CDN в HTML-код на своей странице.</span><span class="sxs-lookup"><span data-stu-id="55b5b-114">Add the CDN reference to the HTML on your page.</span></span>  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css">
    ```   
    
2. <span data-ttu-id="55b5b-115">Используйте значки и шрифты Fabric.</span><span class="sxs-lookup"><span data-stu-id="55b5b-115">Use Fabric icons and fonts.</span></span> 

    <span data-ttu-id="55b5b-p104">Чтобы использовать значок Fabric, разместите элемент "i" на своей странице и сошлитесь на соответствующие классы. Вы можете сами выбирать размер значка, изменяя размер шрифта. Например, в коде ниже показано, как сделать очень большой значок таблицы, который использует цвет themePrimary (#0078d7).</span><span class="sxs-lookup"><span data-stu-id="55b5b-p104">To use a Fabric icon, include the "i" element on your page, and then reference the appropriate classes. You can control the size of the icon by changing the font size. For example, the following code shows how to make an extra-large table icon that uses the themePrimary (#0078d7) color.</span></span> 
   
    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    <span data-ttu-id="55b5b-p105">Чтобы найти другие значки, доступные в Office UI Fabric, используйте функцию поиска на странице [Значки](https://dev.office.com/fabric#/styles/icons). Когда вы найдете значок для надстройки, добавьте к его имени префикс `ms-Icon--`.</span><span class="sxs-lookup"><span data-stu-id="55b5b-p105">To find more icons that are available in Office UI Fabric, use the search feature on the [Icons](https://dev.office.com/fabric#/styles/icons) page. When you find an icon to use in your add-in, be sure to prefix the icon name with `ms-Icon--`.</span></span> 

    <span data-ttu-id="55b5b-121">Сведения о размерах шрифтов и цветах, доступных в Office UI Fabric, см. в разделах [Оформление](https://dev.office.com/fabric#/styles/typography) и [Цвета](https://dev.office.com/fabric#/styles/colors).</span><span class="sxs-lookup"><span data-stu-id="55b5b-121">For information about font sizes and colors that are available in Office UI Fabric, see [Typography](https://dev.office.com/fabric#/styles/typography) and [Colors](https://dev.office.com/fabric#/styles/colors).</span></span>
 
## <a name="use-fabric-components"></a><span data-ttu-id="55b5b-122">Использование компонентов Fabric</span><span class="sxs-lookup"><span data-stu-id="55b5b-122">Use Fabric Components</span></span> 
<span data-ttu-id="55b5b-123">В Fabric есть различные компоненты оформления, которые можно использовать при создании надстроек, в том числе:</span><span class="sxs-lookup"><span data-stu-id="55b5b-123">Fabric provides a variety of UX components that you can use to build your add-in, including the following types of components:</span></span>

- <span data-ttu-id="55b5b-124">Компоненты ввода — например, Button, Checkbox и Toggle.</span><span class="sxs-lookup"><span data-stu-id="55b5b-124">Input components - for example, Button, Checkbox, and Toggle</span></span>
- <span data-ttu-id="55b5b-125">Компоненты навигации — например, сводка и строка навигации</span><span class="sxs-lookup"><span data-stu-id="55b5b-125">Navigation components - for example, Pivot Breadcrumb</span></span>
- <span data-ttu-id="55b5b-126">Компоненты уведомления — например, MessageBar и Callout.</span><span class="sxs-lookup"><span data-stu-id="55b5b-126">Notification components - for example, MessageBar and Callout</span></span>  

<span data-ttu-id="55b5b-127">Не все компоненты Fabric рекомендуются для использования в надстройках. Ниже приведен список компонентов Fabric React UX, которые мы рекомендуем использовать в надстройке:</span><span class="sxs-lookup"><span data-stu-id="55b5b-127">Not all Fabric components are recommended for use in add-ins. Here is a list of Fabric React UX components that we recommend for use in an add-in:</span></span>

- [<span data-ttu-id="55b5b-128">Строка навигации</span><span class="sxs-lookup"><span data-stu-id="55b5b-128">Breadcrumb</span></span>](https://developer.microsoft.com/fabric#/components/breadcrumb)
- [<span data-ttu-id="55b5b-129">Кнопка</span><span class="sxs-lookup"><span data-stu-id="55b5b-129">Button</span></span>](https://developer.microsoft.com/fabric#/components/button)
- [<span data-ttu-id="55b5b-130">Флажок</span><span class="sxs-lookup"><span data-stu-id="55b5b-130">Checkbox</span></span>](https://developer.microsoft.com/fabric#/components/checkbox)
- [<span data-ttu-id="55b5b-131">ChoiceGroup</span><span class="sxs-lookup"><span data-stu-id="55b5b-131">ChoiceGroup</span></span>](https://developer.microsoft.com/fabric#/components/choicegroup)
- [<span data-ttu-id="55b5b-132">Раскрывающееся меню</span><span class="sxs-lookup"><span data-stu-id="55b5b-132">Dropdown</span></span>](https://developer.microsoft.com/fabric#/components/dropdown)
- [<span data-ttu-id="55b5b-133">Подпись</span><span class="sxs-lookup"><span data-stu-id="55b5b-133">Label</span></span>](https://developer.microsoft.com/fabric#/components/label)
- [<span data-ttu-id="55b5b-134">Список</span><span class="sxs-lookup"><span data-stu-id="55b5b-134">List</span></span>](https://developer.microsoft.com/fabric#/components/list)
- [<span data-ttu-id="55b5b-135">Сводка</span><span class="sxs-lookup"><span data-stu-id="55b5b-135">Pivot</span></span>](https://developer.microsoft.com/fabric#/components/pivot)
- [<span data-ttu-id="55b5b-136">TextField</span><span class="sxs-lookup"><span data-stu-id="55b5b-136">TextField</span></span>](https://developer.microsoft.com/fabric#/components/textfield)
- [<span data-ttu-id="55b5b-137">Переключатель</span><span class="sxs-lookup"><span data-stu-id="55b5b-137">Toggle</span></span>](https://developer.microsoft.com/fabric#/components/toggle)

<span data-ttu-id="55b5b-p106">Для создания надстройки можно использовать разные платформы JavaScript, такие как Angular или React. Прежде чем использовать компоненты Fabric со своей платформой, ознакомьтесь с перечисленными ниже ресурсами.</span><span class="sxs-lookup"><span data-stu-id="55b5b-p106">You can use different JavaScript frameworks, such as Angular or React, to build your add-in. To get started using Fabric components with your framework, see the following resources.</span></span>

|<span data-ttu-id="55b5b-140">**Платформа**</span><span class="sxs-lookup"><span data-stu-id="55b5b-140">**Framework**</span></span>|<span data-ttu-id="55b5b-141">**Пример**</span><span class="sxs-lookup"><span data-stu-id="55b5b-141">**Example**</span></span>|
|:------------|:----------|
|<span data-ttu-id="55b5b-142">**React**</span><span class="sxs-lookup"><span data-stu-id="55b5b-142">**React**</span></span>|[<span data-ttu-id="55b5b-143">Использование Office UI Fabric React в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="55b5b-143">Using Office UI Fabric React in Office Add-ins</span></span>](using-office-ui-fabric-react.md )|
|<span data-ttu-id="55b5b-144">**Angular**</span><span class="sxs-lookup"><span data-stu-id="55b5b-144">**Angular**</span></span>| <span data-ttu-id="55b5b-145">См. проект сообщества [ngOfficeUIFabric](http://ngofficeuifabric.com/) с директивами Angular 1.5 и раздел [Обдумайте размещение компонентов Fabric в компонентах Angular 2](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components)</span><span class="sxs-lookup"><span data-stu-id="55b5b-145">See [ngOfficeUIFabric](http://ngofficeuifabric.com/), which is a community project with Angular 1.5 directives, and [Consider wrapping Fabric components with Angular 2 components](../develop/add-ins-with-angular2.md#consider-wrapping-fabric-components-with-angular-components)</span></span>|
