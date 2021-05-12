---
title: Office UI Fabric в надстройках Office
description: Получите обзор использования компонентов Office UI Fabric в Office надстройки.
ms.date: 05/03/2021
localization_priority: Normal
ms.openlocfilehash: 20f926913335197a65ac24e4ec30ed0106b81bae
ms.sourcegitcommit: 132f5082f5bf9500dad0a2eaf89d924c823e575d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330537"
---
# <a name="office-ui-fabric-in-office-add-ins"></a><span data-ttu-id="f0c9b-103">Office UI Fabric в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="f0c9b-103">Office UI Fabric in Office Add-ins</span></span>

<span data-ttu-id="f0c9b-104">Office UI Fabric является интерфейсной платформой JavaScript для создания пользовательских интерфейсов для Office.</span><span class="sxs-lookup"><span data-stu-id="f0c9b-104">Office UI Fabric is a JavaScript front-end framework for building user experiences for Office.</span></span> <span data-ttu-id="f0c9b-105">В Fabric предоставлены компоненты дизайна, которые можно расширять, дорабатывать и использовать в надстройке Office.</span><span class="sxs-lookup"><span data-stu-id="f0c9b-105">Fabric provides visuals-focused components that you can extend, rework, and use in your Office Add-in.</span></span> <span data-ttu-id="f0c9b-106">Так как Fabric использует язык дизайна Office, компоненты дизайна Fabric выглядят в Office очень естественно.</span><span class="sxs-lookup"><span data-stu-id="f0c9b-106">Because Fabric uses the Office Design Language, Fabric's UX components look like a natural extension of Office.</span></span>

<span data-ttu-id="f0c9b-p102">Рекомендуем использовать Office UI Fabric для создания надстроек. Использовать Office UI Fabric необязательно.</span><span class="sxs-lookup"><span data-stu-id="f0c9b-p102">If you are building an add-in, we encourage you to use Office UI Fabric to create your user experience. Using Office UI Fabric is optional.</span></span>

<span data-ttu-id="f0c9b-109">В следующих разделах описано, как начать использовать Fabric для своих потребностей.</span><span class="sxs-lookup"><span data-stu-id="f0c9b-109">The following sections explain how to get started using Fabric to meet your requirements.</span></span>

## <a name="use-fabric-core-icons-fonts-colors"></a><span data-ttu-id="f0c9b-110">Использование Fabric Core: значки, шрифты, цвета</span><span class="sxs-lookup"><span data-stu-id="f0c9b-110">Use Fabric Core: icons, fonts, colors</span></span>

<span data-ttu-id="f0c9b-111">Fabric Core содержит основные элементы языка дизайна, такие как значки, цвета, тип и сетку.</span><span class="sxs-lookup"><span data-stu-id="f0c9b-111">Fabric Core contains basic elements of the design language such as icons, colors, type, and grid.</span></span> <span data-ttu-id="f0c9b-112">Fabric Core не зависит от платформы.</span><span class="sxs-lookup"><span data-stu-id="f0c9b-112">Fabric core is framework independent.</span></span> <span data-ttu-id="f0c9b-113">Fabric Core используется с помощью Fabric React и входит в его состав.</span><span class="sxs-lookup"><span data-stu-id="f0c9b-113">Fabric Core is used by and included with Fabric React.</span></span>

<span data-ttu-id="f0c9b-114">Чтобы начать работу с Fabric Core:</span><span class="sxs-lookup"><span data-stu-id="f0c9b-114">To get started using Fabric Core:</span></span>

1. <span data-ttu-id="f0c9b-115">Добавьте ссылку CDN в HTML-код на своей странице.</span><span class="sxs-lookup"><span data-stu-id="f0c9b-115">Add the CDN reference to the HTML on your page.</span></span>  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css">
    ```

2. <span data-ttu-id="f0c9b-116">Используйте значки и шрифты Fabric.</span><span class="sxs-lookup"><span data-stu-id="f0c9b-116">Use Fabric icons and fonts.</span></span>

    <span data-ttu-id="f0c9b-p104">Чтобы использовать значок Fabric, разместите элемент "i" на своей странице и сошлитесь на соответствующие классы. Вы можете сами выбирать размер значка, изменяя размер шрифта. Например, в коде ниже показано, как сделать очень большой значок таблицы, который использует цвет themePrimary (#0078d7).</span><span class="sxs-lookup"><span data-stu-id="f0c9b-p104">To use a Fabric icon, include the "i" element on your page, and then reference the appropriate classes. You can control the size of the icon by changing the font size. For example, the following code shows how to make an extra-large table icon that uses the themePrimary (#0078d7) color.</span></span>

    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    <span data-ttu-id="f0c9b-p105">Чтобы найти другие значки, доступные в Office UI Fabric, используйте функцию поиска на странице [Значки](https://developer.microsoft.com/fabric#/styles/icons). Когда вы найдете значок для надстройки, добавьте к его имени префикс `ms-Icon--`.</span><span class="sxs-lookup"><span data-stu-id="f0c9b-p105">To find more icons that are available in Office UI Fabric, use the search feature on the [Icons](https://developer.microsoft.com/fabric#/styles/icons) page. When you find an icon to use in your add-in, be sure to prefix the icon name with `ms-Icon--`.</span></span>

    <span data-ttu-id="f0c9b-122">Сведения о размерах шрифтов и цветах, доступных в Office UI Fabric, см. в разделах [Оформление](https://developer.microsoft.com/fabric#/styles/typography) и [Цвета](https://developer.microsoft.com/fabric#/styles/colors).</span><span class="sxs-lookup"><span data-stu-id="f0c9b-122">For information about font sizes and colors that are available in Office UI Fabric, see [Typography](https://developer.microsoft.com/fabric#/styles/typography) and [Colors](https://developer.microsoft.com/fabric#/styles/colors).</span></span>

## <a name="use-fabric-components"></a><span data-ttu-id="f0c9b-123">Использование компонентов Fabric</span><span class="sxs-lookup"><span data-stu-id="f0c9b-123">Use Fabric Components</span></span>

<span data-ttu-id="f0c9b-124">Fabric предоставляет различные компоненты UX, которые можно использовать для создания надстройки.</span><span class="sxs-lookup"><span data-stu-id="f0c9b-124">Fabric provides a variety of UX components that you can use to build your add-in.</span></span> <span data-ttu-id="f0c9b-125">Мы не ожидаем, что все компоненты ткани будут использоваться одной надстройки.</span><span class="sxs-lookup"><span data-stu-id="f0c9b-125">We do not expect that all fabric components will be used by a single add-in.</span></span> <span data-ttu-id="f0c9b-126">Определите оптимальные компоненты для сценария и пользовательского интерфейса (например, может быть трудно правильно отображать [breadcrumb](https://developer.microsoft.com/fabric#/components/breadcrumb) в области задач).</span><span class="sxs-lookup"><span data-stu-id="f0c9b-126">Determine the best components for your scenario and user experience (for example, it may be hard to properly display a [Breadcrumb](https://developer.microsoft.com/fabric#/components/breadcrumb) in the task pane).</span></span>

<span data-ttu-id="f0c9b-127">Ниже приводится список общих компонентов [React Fabric,](https://developer.microsoft.com/fluentui#/controls/web) которые мы рекомендуем использовать в надстройки:</span><span class="sxs-lookup"><span data-stu-id="f0c9b-127">The following is a list of common [Fabric React UX components](https://developer.microsoft.com/fluentui#/controls/web) that we recommend for use in an add-in:</span></span>

- [<span data-ttu-id="f0c9b-128">Кнопка</span><span class="sxs-lookup"><span data-stu-id="f0c9b-128">Button</span></span>](https://developer.microsoft.com/fabric#/components/button)
- [<span data-ttu-id="f0c9b-129">Флажок</span><span class="sxs-lookup"><span data-stu-id="f0c9b-129">Checkbox</span></span>](https://developer.microsoft.com/fabric#/components/checkbox)
- [<span data-ttu-id="f0c9b-130">ChoiceGroup</span><span class="sxs-lookup"><span data-stu-id="f0c9b-130">ChoiceGroup</span></span>](https://developer.microsoft.com/fabric#/components/choicegroup)
- [<span data-ttu-id="f0c9b-131">Раскрывающееся меню</span><span class="sxs-lookup"><span data-stu-id="f0c9b-131">Dropdown</span></span>](https://developer.microsoft.com/fabric#/components/dropdown)
- [<span data-ttu-id="f0c9b-132">Подпись</span><span class="sxs-lookup"><span data-stu-id="f0c9b-132">Label</span></span>](https://developer.microsoft.com/fabric#/components/label)
- [<span data-ttu-id="f0c9b-133">Список</span><span class="sxs-lookup"><span data-stu-id="f0c9b-133">List</span></span>](https://developer.microsoft.com/fabric#/components/list)
- [<span data-ttu-id="f0c9b-134">Сводка</span><span class="sxs-lookup"><span data-stu-id="f0c9b-134">Pivot</span></span>](https://developer.microsoft.com/fabric#/components/pivot)
- [<span data-ttu-id="f0c9b-135">TextField</span><span class="sxs-lookup"><span data-stu-id="f0c9b-135">TextField</span></span>](https://developer.microsoft.com/fabric#/components/textfield)
- [<span data-ttu-id="f0c9b-136">Переключатель</span><span class="sxs-lookup"><span data-stu-id="f0c9b-136">Toggle</span></span>](https://developer.microsoft.com/fabric#/components/toggle)

<span data-ttu-id="f0c9b-p107">Для создания надстройки можно использовать разные платформы JavaScript, такие как Angular или React. Прежде чем использовать компоненты Fabric со своей платформой, ознакомьтесь с перечисленными ниже ресурсами.</span><span class="sxs-lookup"><span data-stu-id="f0c9b-p107">You can use different JavaScript frameworks, such as Angular or React, to build your add-in. To get started using Fabric components with your framework, see the following resources.</span></span>

|<span data-ttu-id="f0c9b-139">**Платформа**</span><span class="sxs-lookup"><span data-stu-id="f0c9b-139">**Framework**</span></span>|<span data-ttu-id="f0c9b-140">**Пример**</span><span class="sxs-lookup"><span data-stu-id="f0c9b-140">**Example**</span></span>|
|:------------|:----------|
|<span data-ttu-id="f0c9b-141">**React**</span><span class="sxs-lookup"><span data-stu-id="f0c9b-141">**React**</span></span>|[<span data-ttu-id="f0c9b-142">Использование Office UI Fabric React в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="f0c9b-142">Using Office UI Fabric React in Office Add-ins</span></span>](using-office-ui-fabric-react.md )|
