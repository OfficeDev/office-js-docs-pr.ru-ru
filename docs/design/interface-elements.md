---
title: Элементы пользовательского интерфейса Office для надстроек Office
description: Получите обзор различных элементов пользовательского интерфейса в Office надстройки.
ms.date: 12/24/2019
localization_priority: Normal
ms.openlocfilehash: 5d0a1576d850f2291c28e6bb39554cbb0403f50b
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076331"
---
# <a name="office-ui-elements-for-office-add-ins"></a><span data-ttu-id="5a8ab-103">Элементы пользовательского интерфейса Office для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="5a8ab-103">Office UI elements for Office Add-ins</span></span>

<span data-ttu-id="5a8ab-p101">Для расширения пользовательского интерфейса Office, в том числе команд надстроек и контейнеров HTML, можно использовать несколько типов элементов пользовательского интерфейса, которые полностью совместимы с Office и рядом платформ. Вы можете вставить пользовательский веб-код в любой из этих элементов.</span><span class="sxs-lookup"><span data-stu-id="5a8ab-p101">You can use several types of UI elements to extend the Office UI, including add-in commands and HTML containers. These UI elements look like a natural extension of Office and work across platforms. You can insert your custom web-based code into any of these elements.</span></span>

<span data-ttu-id="5a8ab-107">На рисунке ниже приведены типы элементов пользовательского интерфейса Office, которые можно создать.</span><span class="sxs-lookup"><span data-stu-id="5a8ab-107">The following image shows the types of Office UI elements that you can create.</span></span>

![Схема, показывающая команды надстройки в ленте, области задач и диалоговом окне/надстройке контента в Office документе.](../images/add-in-ui-elements.png)

## <a name="add-in-commands"></a><span data-ttu-id="5a8ab-109">Команды надстроек</span><span class="sxs-lookup"><span data-stu-id="5a8ab-109">Add-in commands</span></span>

<span data-ttu-id="5a8ab-110">Используйте [команды надстройки,](add-in-commands.md) чтобы добавить точки входа в надстройки к ленте Приложение Office.</span><span class="sxs-lookup"><span data-stu-id="5a8ab-110">Use [add-in commands](add-in-commands.md) to add entry points to your add-in to the Office app ribbon.</span></span> <span data-ttu-id="5a8ab-111">Команды запускают действия в надстройке путем выполнения кода JavaScript или запуска контейнера HTML.</span><span class="sxs-lookup"><span data-stu-id="5a8ab-111">Commands start actions in your add-in either by running JavaScript code, or by launching an HTML container.</span></span> <span data-ttu-id="5a8ab-112">Можно создать два типа команд надстроек.</span><span class="sxs-lookup"><span data-stu-id="5a8ab-112">You can create two types of add-in commands.</span></span>

|<span data-ttu-id="5a8ab-113">Тип команды</span><span class="sxs-lookup"><span data-stu-id="5a8ab-113">Command type</span></span>|<span data-ttu-id="5a8ab-114">Описание</span><span class="sxs-lookup"><span data-stu-id="5a8ab-114">Description</span></span>|
|:---------------|:--------------|
|<span data-ttu-id="5a8ab-115">Кнопки, меню и вкладки на ленте</span><span class="sxs-lookup"><span data-stu-id="5a8ab-115">Ribbon buttons, menus, and tabs</span></span>|<span data-ttu-id="5a8ab-p103">Позволяют добавлять в Office пользовательские кнопки, меню (раскрывающиеся меню) или вкладки на ленте по умолчанию. Кнопки и меню используются для запуска действия в Office. Вкладки позволяют сгруппировать и упорядочить кнопки и меню.</span><span class="sxs-lookup"><span data-stu-id="5a8ab-p103">Use to add custom buttons, menus (dropdowns), or tabs to the default ribbon in Office. Use Buttons and menus to trigger an action in Office. Use tabs to group and organize buttons and menus.</span></span>|
|<span data-ttu-id="5a8ab-119">Контекстные меню</span><span class="sxs-lookup"><span data-stu-id="5a8ab-119">Context menus</span></span>| <span data-ttu-id="5a8ab-p104">Используются для расширения контекстного меню по умолчанию. Контекстные меню отображаются, когда пользователи щелкают правой кнопкой мыши текст в документе Office или таблице Excel.</span><span class="sxs-lookup"><span data-stu-id="5a8ab-p104">Use to extend the default context menu. Context menus are displayed when users right-click text in an Office document or a table in Excel.</span></span>|

## <a name="html-containers"></a><span data-ttu-id="5a8ab-122">Контейнеры HTML</span><span class="sxs-lookup"><span data-stu-id="5a8ab-122">HTML containers</span></span>

<span data-ttu-id="5a8ab-p105">Контейнеры HTML позволяют внедрить код пользовательского интерфейса на основе HTML в клиентах Office. Эти веб-страницы затем могут ссылаться на API JavaScript для Office для взаимодействия с содержимым в документе. Можно создать HTML-контейнеры трех типов.</span><span class="sxs-lookup"><span data-stu-id="5a8ab-p105">Use HTML containers to embed HTML-based UI code within Office clients. These web pages can then reference the Office JavaScript API to interact with content in the document. You can create three types of HTML containers.</span></span>

|<span data-ttu-id="5a8ab-126">Контейнер HTML</span><span class="sxs-lookup"><span data-stu-id="5a8ab-126">HTML container</span></span>|<span data-ttu-id="5a8ab-127">Описание</span><span class="sxs-lookup"><span data-stu-id="5a8ab-127">Description</span></span>|
|:-----------------|:--------------|
|[<span data-ttu-id="5a8ab-128">Области задач</span><span class="sxs-lookup"><span data-stu-id="5a8ab-128">Task panes</span></span>](task-pane-add-ins.md)|<span data-ttu-id="5a8ab-p106">Отображение собственного пользовательского интерфейса в правой части документа Office. Области задач позволяют пользователям взаимодействовать с вашей надстройкой, работая с документом Office.</span><span class="sxs-lookup"><span data-stu-id="5a8ab-p106">Display custom UI in the right pane of the Office document. Use task panes to allow users to interact with your add-in side-by-side with the Office document.</span></span>|
|[<span data-ttu-id="5a8ab-131">Контентные надстройки</span><span class="sxs-lookup"><span data-stu-id="5a8ab-131">Content add-ins</span></span>](content-add-ins.md)|<span data-ttu-id="5a8ab-p107">Отображение пользовательского интерфейса, внедренного в документы Office. Контентные надстройки позволяют пользователям взаимодействовать с вашей надстройкой непосредственно в документе Office. Например, может понадобиться отобразить внешнее содержимое (видео или визуализации данных из других источников).</span><span class="sxs-lookup"><span data-stu-id="5a8ab-p107">Display custom UI embedded within Office documents. Use content add-ins to allow users to interact with your add-in directly within the Office document. For example, you might want to show external content such as videos or data visualizations from other sources.</span></span> |
|[<span data-ttu-id="5a8ab-135">Диалоговые окна</span><span class="sxs-lookup"><span data-stu-id="5a8ab-135">Dialog boxes</span></span>](dialog-boxes.md)|<span data-ttu-id="5a8ab-p108">Отображение пользовательского интерфейса в диалоговом окне, которое накладывается на документ Office. Используйте диалоговое окно для действий, которые требуют внимания и не требуют непосредственного взаимодействия с документом.</span><span class="sxs-lookup"><span data-stu-id="5a8ab-p108">Display custom UI in a dialog box that overlays the Office document. Use a dialog box for interactions that require focus and more real estate, and do not require a side-by-side interaction with the document.</span></span>|

## <a name="see-also"></a><span data-ttu-id="5a8ab-138">См. также</span><span class="sxs-lookup"><span data-stu-id="5a8ab-138">See also</span></span>

- [<span data-ttu-id="5a8ab-139">Команды надстроек для Excel, Word и PowerPoint</span><span class="sxs-lookup"><span data-stu-id="5a8ab-139">Add-in commands for Excel, Word, and PowerPoint</span></span>](add-in-commands.md)
- [<span data-ttu-id="5a8ab-140">Области задач</span><span class="sxs-lookup"><span data-stu-id="5a8ab-140">Task panes</span></span>](task-pane-add-ins.md)
- [<span data-ttu-id="5a8ab-141">Контентные надстройки</span><span class="sxs-lookup"><span data-stu-id="5a8ab-141">Content add-ins</span></span>](content-add-ins.md)
- [<span data-ttu-id="5a8ab-142">Диалоговые окна</span><span class="sxs-lookup"><span data-stu-id="5a8ab-142">Dialog boxes</span></span>](dialog-boxes.md)
