---
title: Элементы пользовательского интерфейса Office для надстроек Office
description: ''
ms.date: 12/24/2019
localization_priority: Priority
ms.openlocfilehash: 1f680d93c995d306717c7b5a81af073f378b0fbd
ms.sourcegitcommit: 350f5c6954dec3e9384e2030cd3265aaba7ae904
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/23/2019
ms.locfileid: "40851244"
---
# <a name="office-ui-elements-for-office-add-ins"></a><span data-ttu-id="97d6d-102">Элементы пользовательского интерфейса Office для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="97d6d-102">Office UI elements for Office Add-ins</span></span>

<span data-ttu-id="97d6d-p101">Для расширения пользовательского интерфейса Office, в том числе команд надстроек и контейнеров HTML, можно использовать несколько типов элементов пользовательского интерфейса, которые полностью совместимы с Office и рядом платформ. Вы можете вставить пользовательский веб-код в любой из этих элементов.</span><span class="sxs-lookup"><span data-stu-id="97d6d-p101">You can use several types of UI elements to extend the Office UI, including add-in commands and HTML containers. These UI elements look like a natural extension of Office and work across platforms. You can insert your custom web-based code into any of these elements.</span></span>

<span data-ttu-id="97d6d-106">На рисунке ниже приведены типы элементов пользовательского интерфейса Office, которые можно создать.</span><span class="sxs-lookup"><span data-stu-id="97d6d-106">The following image shows the types of Office UI elements that you can create.</span></span>

![Изображение с командами надстроек на ленте, областью задач и диалоговым окном в документе Office](../images/add-in-ui-elements.png)

## <a name="add-in-commands"></a><span data-ttu-id="97d6d-108">Команды надстроек</span><span class="sxs-lookup"><span data-stu-id="97d6d-108">Add-in commands</span></span>

<span data-ttu-id="97d6d-p102">[Команды надстроек](add-in-commands.md) используются для добавления точек входа в надстройку на ленте Office. Команды запускают действия в надстройке путем выполнения кода JavaScript или запуска контейнера HTML. Можно создать два типа команд надстроек.</span><span class="sxs-lookup"><span data-stu-id="97d6d-p102">Use [add-in commands](add-in-commands.md) to add entry points to your add-in to the Office ribbon. Commands start actions in your add-in either by running JavaScript code, or by launching an HTML container. You can create two types of add-in commands.</span></span>

|<span data-ttu-id="97d6d-112">**Тип команды**</span><span class="sxs-lookup"><span data-stu-id="97d6d-112">**Command type**</span></span>|<span data-ttu-id="97d6d-113">**Описание**</span><span class="sxs-lookup"><span data-stu-id="97d6d-113">**Description**</span></span>|
|:---------------|:--------------|
|<span data-ttu-id="97d6d-114">Кнопки, меню и вкладки на ленте</span><span class="sxs-lookup"><span data-stu-id="97d6d-114">Ribbon buttons, menus, and tabs</span></span>|<span data-ttu-id="97d6d-p103">Позволяют добавлять в Office пользовательские кнопки, меню (раскрывающиеся меню) или вкладки на ленте по умолчанию. Кнопки и меню используются для запуска действия в Office. Вкладки позволяют сгруппировать и упорядочить кнопки и меню.</span><span class="sxs-lookup"><span data-stu-id="97d6d-p103">Use to add custom buttons, menus (dropdowns), or tabs to the default ribbon in Office. Use Buttons and menus to trigger an action in Office. Use tabs to group and organize buttons and menus.</span></span>|
|<span data-ttu-id="97d6d-118">Контекстные меню</span><span class="sxs-lookup"><span data-stu-id="97d6d-118">Context menus</span></span>| <span data-ttu-id="97d6d-p104">Используются для расширения контекстного меню по умолчанию. Контекстные меню отображаются, когда пользователи щелкают правой кнопкой мыши текст в документе Office или таблице Excel.</span><span class="sxs-lookup"><span data-stu-id="97d6d-p104">Use to extend the default context menu. Context menus are displayed when users right-click text in an Office document or a table in Excel.</span></span>| 

## <a name="html-containers"></a><span data-ttu-id="97d6d-121">Контейнеры HTML</span><span class="sxs-lookup"><span data-stu-id="97d6d-121">HTML containers</span></span>

<span data-ttu-id="97d6d-p105">Контейнеры HTML позволяют внедрить код пользовательского интерфейса на основе HTML в клиентах Office. Эти веб-страницы затем могут ссылаться на API JavaScript для Office для взаимодействия с содержимым в документе. Можно создать HTML-контейнеры трех типов.</span><span class="sxs-lookup"><span data-stu-id="97d6d-p105">Use HTML containers to embed HTML-based UI code within Office clients. These web pages can then reference the Office JavaScript API to interact with content in the document. You can create three types of HTML containers.</span></span>

|<span data-ttu-id="97d6d-125">**Контейнер HTML**</span><span class="sxs-lookup"><span data-stu-id="97d6d-125">**HTML container**</span></span>|<span data-ttu-id="97d6d-126">**Описание**</span><span class="sxs-lookup"><span data-stu-id="97d6d-126">**Description**</span></span>|
|:-----------------|:--------------|
|[<span data-ttu-id="97d6d-127">Области задач</span><span class="sxs-lookup"><span data-stu-id="97d6d-127">Task panes</span></span>](task-pane-add-ins.md)|<span data-ttu-id="97d6d-p106">Отображение собственного пользовательского интерфейса в правой части документа Office. Области задач позволяют пользователям взаимодействовать с вашей надстройкой, работая с документом Office.</span><span class="sxs-lookup"><span data-stu-id="97d6d-p106">Display custom UI in the right pane of the Office document. Use task panes to allow users to interact with your add-in side-by-side with the Office document.</span></span>|
|[<span data-ttu-id="97d6d-130">Контентные надстройки</span><span class="sxs-lookup"><span data-stu-id="97d6d-130">Content add-ins</span></span>](content-add-ins.md)|<span data-ttu-id="97d6d-p107">Отображение пользовательского интерфейса, внедренного в документы Office. Контентные надстройки позволяют пользователям взаимодействовать с вашей надстройкой непосредственно в документе Office. Например, может понадобиться отобразить внешнее содержимое (видео или визуализации данных из других источников).</span><span class="sxs-lookup"><span data-stu-id="97d6d-p107">Display custom UI embedded within Office documents. Use content add-ins to allow users to interact with your add-in directly within the Office document. For example, you might want to show external content such as videos or data visualizations from other sources.</span></span> |
|[<span data-ttu-id="97d6d-134">Диалоговые окна</span><span class="sxs-lookup"><span data-stu-id="97d6d-134">Dialog boxes</span></span>](dialog-boxes.md)|<span data-ttu-id="97d6d-p108">Отображение пользовательского интерфейса в диалоговом окне, которое накладывается на документ Office. Используйте диалоговое окно для действий, которые требуют внимания и не требуют непосредственного взаимодействия с документом.</span><span class="sxs-lookup"><span data-stu-id="97d6d-p108">Display custom UI in a dialog box that overlays the Office document. Use a dialog box for interactions that require focus and more real estate, and do not require a side-by-side interaction with the document.</span></span>|

## <a name="see-also"></a><span data-ttu-id="97d6d-137">См. также</span><span class="sxs-lookup"><span data-stu-id="97d6d-137">See also</span></span>

- [<span data-ttu-id="97d6d-138">Команды надстроек для Excel, Word и PowerPoint</span><span class="sxs-lookup"><span data-stu-id="97d6d-138">Add-in commands for Excel, Word, and PowerPoint</span></span>](add-in-commands.md)
- [<span data-ttu-id="97d6d-139">Области задач</span><span class="sxs-lookup"><span data-stu-id="97d6d-139">Task panes</span></span>](task-pane-add-ins.md)
- [<span data-ttu-id="97d6d-140">Контентные надстройки</span><span class="sxs-lookup"><span data-stu-id="97d6d-140">Content add-ins</span></span>](content-add-ins.md)
- [<span data-ttu-id="97d6d-141">Диалоговые окна</span><span class="sxs-lookup"><span data-stu-id="97d6d-141">Dialog boxes</span></span>](dialog-boxes.md)
