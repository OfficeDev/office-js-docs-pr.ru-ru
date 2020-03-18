---
title: Диалоговые окна в надстройках Office
description: Ознакомьтесь с рекомендациями по визуальному дизайну диалоговых окон в надстройках Office.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 2f3b25fac7f12494e6b5a1e0a32e72baa345e978
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717196"
---
# <a name="dialog-boxes-in-office-add-ins"></a><span data-ttu-id="e0688-103">Диалоговые окна в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="e0688-103">Dialog boxes in Office Add-ins</span></span>
 
<span data-ttu-id="e0688-p101">Диалоговые окна — окна, которые накладываются на активное окно приложения Office. Вы можете использовать диалоговые окна, чтобы показывать страницы входа, которые нельзя открыть непосредственно в области задач, запросы на подтверждение действий, предпринятых пользователем, или видео, которые будут слишком маленькими в области задач.</span><span class="sxs-lookup"><span data-stu-id="e0688-p101">Dialog boxes are surfaces that float above the active Office application window. You can use dialog boxes to provide additional screen space for tasks such as sign-in pages that can't be opened directly in a task pane or requests to confirm an action taken by a user, or to show videos that might be too small if confined to a task pane.</span></span>

<span data-ttu-id="e0688-106">*Рисунок 1. Типичный макет диалогового окна*</span><span class="sxs-lookup"><span data-stu-id="e0688-106">*Figure 1. Typical layout for a dialog box*</span></span>

![Изображение, на котором показан типичный макет диалогового окна](../images/overview-with-app-dialog.png)

## <a name="best-practices"></a><span data-ttu-id="e0688-108">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="e0688-108">Best practices</span></span>

|<span data-ttu-id="e0688-109">**Рекомендуется**</span><span class="sxs-lookup"><span data-stu-id="e0688-109">**Do**</span></span>|<span data-ttu-id="e0688-110">**Не рекомендуется**</span><span class="sxs-lookup"><span data-stu-id="e0688-110">**Don't**</span></span>|
|:-----|:--------|
|<ul><li><span data-ttu-id="e0688-111">Укажите описательное название, содержащее имя надстройки и название текущей задачи.</span><span class="sxs-lookup"><span data-stu-id="e0688-111">Include a descriptive title that includes your add-in name along with the current task.</span></span></li></ul>|<ul><li><span data-ttu-id="e0688-112">Не включайте в него название вашей компании.</span><span class="sxs-lookup"><span data-stu-id="e0688-112">Don't append your company name to the title.</span></span></li></ul>|
||<ul><li><span data-ttu-id="e0688-113">Не открывайте диалоговое окно, если этого не требует сценарий.</span><span class="sxs-lookup"><span data-stu-id="e0688-113">Don't open a dialog box unless the scenario requires it.</span></span></li></ul>|

## <a name="implementation"></a><span data-ttu-id="e0688-114">Реализация</span><span class="sxs-lookup"><span data-stu-id="e0688-114">Implementation</span></span>

<span data-ttu-id="e0688-115">Пример реализации диалогового окна с использованием Dialog API для надстроек Office см. в [этой статье](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) на сайте GitHub.</span><span class="sxs-lookup"><span data-stu-id="e0688-115">For a sample that implements a dialog box, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) in GitHub.</span></span>

## <a name="see-also"></a><span data-ttu-id="e0688-116">См. также</span><span class="sxs-lookup"><span data-stu-id="e0688-116">See also</span></span>

- [<span data-ttu-id="e0688-117">Объект Dialog</span><span class="sxs-lookup"><span data-stu-id="e0688-117">Dialog object</span></span>](/javascript/api/office/office.dialog)
- [<span data-ttu-id="e0688-118">Конструктивные шаблоны для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="e0688-118">UX design patterns for Office Add-ins</span></span>](../design/ux-design-pattern-templates.md)
