---
title: Элементы пользовательского интерфейса Office для надстроек Office
description: Получите обзор различных типов элементов пользовательского интерфейса в надстройке Office.
ms.date: 12/24/2019
localization_priority: Normal
ms.openlocfilehash: 3e5ff84cb0d8417d6fab5ec6a39575ce7ff74e23
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132048"
---
# <a name="office-ui-elements-for-office-add-ins"></a><span data-ttu-id="f0faa-103">Элементы пользовательского интерфейса Office для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="f0faa-103">Office UI elements for Office Add-ins</span></span>

<span data-ttu-id="f0faa-p101">Для расширения пользовательского интерфейса Office, в том числе команд надстроек и контейнеров HTML, можно использовать несколько типов элементов пользовательского интерфейса, которые полностью совместимы с Office и рядом платформ. Вы можете вставить пользовательский веб-код в любой из этих элементов.</span><span class="sxs-lookup"><span data-stu-id="f0faa-p101">You can use several types of UI elements to extend the Office UI, including add-in commands and HTML containers. These UI elements look like a natural extension of Office and work across platforms. You can insert your custom web-based code into any of these elements.</span></span>

<span data-ttu-id="f0faa-107">На рисунке ниже приведены типы элементов пользовательского интерфейса Office, которые можно создать.</span><span class="sxs-lookup"><span data-stu-id="f0faa-107">The following image shows the types of Office UI elements that you can create.</span></span>

![Диаграмма, на которой показаны команды надстройки на ленте, области задач и надстройки диалогового окна и контентных надстроек в документе Office](../images/add-in-ui-elements.png)

## <a name="add-in-commands"></a><span data-ttu-id="f0faa-109">Команды надстроек</span><span class="sxs-lookup"><span data-stu-id="f0faa-109">Add-in commands</span></span>

<span data-ttu-id="f0faa-110">Используйте [команды надстройки](add-in-commands.md) для добавления точек входа в надстройку на ленту приложения Office.</span><span class="sxs-lookup"><span data-stu-id="f0faa-110">Use [add-in commands](add-in-commands.md) to add entry points to your add-in to the Office app ribbon.</span></span> <span data-ttu-id="f0faa-111">Команды запускают действия в надстройке путем выполнения кода JavaScript или запуска контейнера HTML.</span><span class="sxs-lookup"><span data-stu-id="f0faa-111">Commands start actions in your add-in either by running JavaScript code, or by launching an HTML container.</span></span> <span data-ttu-id="f0faa-112">Можно создать два типа команд надстроек.</span><span class="sxs-lookup"><span data-stu-id="f0faa-112">You can create two types of add-in commands.</span></span>

|<span data-ttu-id="f0faa-113">Тип команды</span><span class="sxs-lookup"><span data-stu-id="f0faa-113">Command type</span></span>|<span data-ttu-id="f0faa-114">Описание</span><span class="sxs-lookup"><span data-stu-id="f0faa-114">Description</span></span>|
|:---------------|:--------------|
|<span data-ttu-id="f0faa-115">Кнопки, меню и вкладки на ленте</span><span class="sxs-lookup"><span data-stu-id="f0faa-115">Ribbon buttons, menus, and tabs</span></span>|<span data-ttu-id="f0faa-p103">Позволяют добавлять в Office пользовательские кнопки, меню (раскрывающиеся меню) или вкладки на ленте по умолчанию. Кнопки и меню используются для запуска действия в Office. Вкладки позволяют сгруппировать и упорядочить кнопки и меню.</span><span class="sxs-lookup"><span data-stu-id="f0faa-p103">Use to add custom buttons, menus (dropdowns), or tabs to the default ribbon in Office. Use Buttons and menus to trigger an action in Office. Use tabs to group and organize buttons and menus.</span></span>|
|<span data-ttu-id="f0faa-119">Контекстные меню</span><span class="sxs-lookup"><span data-stu-id="f0faa-119">Context menus</span></span>| <span data-ttu-id="f0faa-p104">Используются для расширения контекстного меню по умолчанию. Контекстные меню отображаются, когда пользователи щелкают правой кнопкой мыши текст в документе Office или таблице Excel.</span><span class="sxs-lookup"><span data-stu-id="f0faa-p104">Use to extend the default context menu. Context menus are displayed when users right-click text in an Office document or a table in Excel.</span></span>|

## <a name="html-containers"></a><span data-ttu-id="f0faa-122">Контейнеры HTML</span><span class="sxs-lookup"><span data-stu-id="f0faa-122">HTML containers</span></span>

<span data-ttu-id="f0faa-p105">Контейнеры HTML позволяют внедрить код пользовательского интерфейса на основе HTML в клиентах Office. Эти веб-страницы затем могут ссылаться на API JavaScript для Office для взаимодействия с содержимым в документе. Можно создать HTML-контейнеры трех типов.</span><span class="sxs-lookup"><span data-stu-id="f0faa-p105">Use HTML containers to embed HTML-based UI code within Office clients. These web pages can then reference the Office JavaScript API to interact with content in the document. You can create three types of HTML containers.</span></span>

|<span data-ttu-id="f0faa-126">Контейнер HTML</span><span class="sxs-lookup"><span data-stu-id="f0faa-126">HTML container</span></span>|<span data-ttu-id="f0faa-127">Описание</span><span class="sxs-lookup"><span data-stu-id="f0faa-127">Description</span></span>|
|:-----------------|:--------------|
|[<span data-ttu-id="f0faa-128">Области задач</span><span class="sxs-lookup"><span data-stu-id="f0faa-128">Task panes</span></span>](task-pane-add-ins.md)|<span data-ttu-id="f0faa-p106">Отображение собственного пользовательского интерфейса в правой части документа Office. Области задач позволяют пользователям взаимодействовать с вашей надстройкой, работая с документом Office.</span><span class="sxs-lookup"><span data-stu-id="f0faa-p106">Display custom UI in the right pane of the Office document. Use task panes to allow users to interact with your add-in side-by-side with the Office document.</span></span>|
|[<span data-ttu-id="f0faa-131">Контентные надстройки</span><span class="sxs-lookup"><span data-stu-id="f0faa-131">Content add-ins</span></span>](content-add-ins.md)|<span data-ttu-id="f0faa-p107">Отображение пользовательского интерфейса, внедренного в документы Office. Контентные надстройки позволяют пользователям взаимодействовать с вашей надстройкой непосредственно в документе Office. Например, может понадобиться отобразить внешнее содержимое (видео или визуализации данных из других источников).</span><span class="sxs-lookup"><span data-stu-id="f0faa-p107">Display custom UI embedded within Office documents. Use content add-ins to allow users to interact with your add-in directly within the Office document. For example, you might want to show external content such as videos or data visualizations from other sources.</span></span> |
|[<span data-ttu-id="f0faa-135">Диалоговые окна</span><span class="sxs-lookup"><span data-stu-id="f0faa-135">Dialog boxes</span></span>](dialog-boxes.md)|<span data-ttu-id="f0faa-p108">Отображение пользовательского интерфейса в диалоговом окне, которое накладывается на документ Office. Используйте диалоговое окно для действий, которые требуют внимания и не требуют непосредственного взаимодействия с документом.</span><span class="sxs-lookup"><span data-stu-id="f0faa-p108">Display custom UI in a dialog box that overlays the Office document. Use a dialog box for interactions that require focus and more real estate, and do not require a side-by-side interaction with the document.</span></span>|

## <a name="see-also"></a><span data-ttu-id="f0faa-138">См. также</span><span class="sxs-lookup"><span data-stu-id="f0faa-138">See also</span></span>

- [<span data-ttu-id="f0faa-139">Команды надстроек для Excel, Word и PowerPoint</span><span class="sxs-lookup"><span data-stu-id="f0faa-139">Add-in commands for Excel, Word, and PowerPoint</span></span>](add-in-commands.md)
- [<span data-ttu-id="f0faa-140">Области задач</span><span class="sxs-lookup"><span data-stu-id="f0faa-140">Task panes</span></span>](task-pane-add-ins.md)
- [<span data-ttu-id="f0faa-141">Контентные надстройки</span><span class="sxs-lookup"><span data-stu-id="f0faa-141">Content add-ins</span></span>](content-add-ins.md)
- [<span data-ttu-id="f0faa-142">Диалоговые окна</span><span class="sxs-lookup"><span data-stu-id="f0faa-142">Dialog boxes</span></span>](dialog-boxes.md)
