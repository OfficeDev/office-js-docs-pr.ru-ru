---
title: Диалоговые окна в надстройках Office
description: Ознакомьтесь с рекомендациями по визуальному дизайну диалоговых окон в надстройках Office.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: ab8ca2e768c63a53b05ed2d9ef459050455231fb
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132055"
---
# <a name="dialog-boxes-in-office-add-ins"></a><span data-ttu-id="bedf9-103">Диалоговые окна в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="bedf9-103">Dialog boxes in Office Add-ins</span></span>

<span data-ttu-id="bedf9-p101">Диалоговые окна — окна, которые накладываются на активное окно приложения Office. Вы можете использовать диалоговые окна, чтобы показывать страницы входа, которые нельзя открыть непосредственно в области задач, запросы на подтверждение действий, предпринятых пользователем, или видео, которые будут слишком маленькими в области задач.</span><span class="sxs-lookup"><span data-stu-id="bedf9-p101">Dialog boxes are surfaces that float above the active Office application window. You can use dialog boxes to provide additional screen space for tasks such as sign-in pages that can't be opened directly in a task pane or requests to confirm an action taken by a user, or to show videos that might be too small if confined to a task pane.</span></span>

<span data-ttu-id="bedf9-106">*Рисунок 1. Типичный макет диалогового окна*</span><span class="sxs-lookup"><span data-stu-id="bedf9-106">*Figure 1. Typical layout for a dialog box*</span></span>

![Типичный макет диалогового окна, отображаемого в приложении Office](../images/overview-with-app-dialog.png)

## <a name="best-practices"></a><span data-ttu-id="bedf9-108">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="bedf9-108">Best practices</span></span>

|<span data-ttu-id="bedf9-109">Правильно</span><span class="sxs-lookup"><span data-stu-id="bedf9-109">Do</span></span>|<span data-ttu-id="bedf9-110">Неправильно</span><span class="sxs-lookup"><span data-stu-id="bedf9-110">Don't</span></span>|
|:-----|:--------|
|<ul><li><span data-ttu-id="bedf9-111">Укажите описательное название, содержащее имя надстройки и название текущей задачи.</span><span class="sxs-lookup"><span data-stu-id="bedf9-111">Include a descriptive title that includes your add-in name along with the current task.</span></span></li></ul>|<ul><li><span data-ttu-id="bedf9-112">Не включайте в него название вашей компании.</span><span class="sxs-lookup"><span data-stu-id="bedf9-112">Don't append your company name to the title.</span></span></li></ul>|
||<ul><li><span data-ttu-id="bedf9-113">Не открывайте диалоговое окно, если этого не требует сценарий.</span><span class="sxs-lookup"><span data-stu-id="bedf9-113">Don't open a dialog box unless the scenario requires it.</span></span></li></ul>|

## <a name="implementation"></a><span data-ttu-id="bedf9-114">Реализация</span><span class="sxs-lookup"><span data-stu-id="bedf9-114">Implementation</span></span>

<span data-ttu-id="bedf9-115">Пример реализации диалогового окна с использованием Dialog API для надстроек Office см. в [этой статье](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) на сайте GitHub.</span><span class="sxs-lookup"><span data-stu-id="bedf9-115">For a sample that implements a dialog box, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) in GitHub.</span></span>

## <a name="see-also"></a><span data-ttu-id="bedf9-116">См. также</span><span class="sxs-lookup"><span data-stu-id="bedf9-116">See also</span></span>

- [<span data-ttu-id="bedf9-117">Объект Dialog</span><span class="sxs-lookup"><span data-stu-id="bedf9-117">Dialog object</span></span>](/javascript/api/office/office.dialog)
- [<span data-ttu-id="bedf9-118">Конструктивные шаблоны для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="bedf9-118">UX design patterns for Office Add-ins</span></span>](../design/ux-design-pattern-templates.md)
