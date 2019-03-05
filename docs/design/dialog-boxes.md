---
title: Диалоговые окна в надстройках Office
description: ''
ms.date: 2/28/2019
localization_priority: Priority
ms.openlocfilehash: 1710d609910cc3c15143605570f97d013a104194
ms.sourcegitcommit: f7f3d38ae4430e2218bf0abe7bb2976108de3579
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/01/2019
ms.locfileid: "30359221"
---
# <a name="dialog-boxes-in-office-add-ins"></a><span data-ttu-id="0745d-102">Диалоговые окна в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="0745d-102">Dialog boxes in Office Add-ins</span></span>
 
<span data-ttu-id="0745d-p101">Диалоговые окна — окна, которые накладываются на активное окно приложения Office. Вы можете использовать диалоговые окна, чтобы показывать страницы входа, которые нельзя открыть непосредственно в области задач, запросы на подтверждение действий, предпринятых пользователем, или видео, которые будут слишком маленькими в области задач.</span><span class="sxs-lookup"><span data-stu-id="0745d-p101">Dialog boxes are surfaces that float above the active Office application window. You can use dialog boxes to provide additional screen space for tasks such as sign-in pages that can't be opened directly in a task pane or requests to confirm an action taken by a user, or to show videos that might be too small if confined to a task pane.</span></span>

<span data-ttu-id="0745d-105">*Рисунок 1. Типичный макет диалогового окна*</span><span class="sxs-lookup"><span data-stu-id="0745d-105">*Figure 1. Typical layout for a dialog box*</span></span>

![Изображение, на котором показан типичный макет диалогового окна](../images/overview-with-app-dialog.png)

## <a name="best-practices"></a><span data-ttu-id="0745d-107">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="0745d-107">Best practices</span></span>

|<span data-ttu-id="0745d-108">**Рекомендуется**</span><span class="sxs-lookup"><span data-stu-id="0745d-108">**Do**</span></span>|<span data-ttu-id="0745d-109">**Не рекомендуется**</span><span class="sxs-lookup"><span data-stu-id="0745d-109">**Don't**</span></span>|
|:-----|:--------|
|<ul><li><span data-ttu-id="0745d-110">Укажите описательное название, содержащее имя надстройки и название текущей задачи.</span><span class="sxs-lookup"><span data-stu-id="0745d-110">Include a descriptive title that includes your add-in name along with the current task.</span></span></li></ul>|<ul><li><span data-ttu-id="0745d-111">Не включайте в него название вашей компании.</span><span class="sxs-lookup"><span data-stu-id="0745d-111">Don't append your company name to the title.</span></span></li></ul>|
||<ul><li><span data-ttu-id="0745d-112">Не открывайте диалоговое окно, если этого не требует сценарий.</span><span class="sxs-lookup"><span data-stu-id="0745d-112">Don't open a dialog box unless the scenario requires it.</span></span></li></ul>|

## <a name="implementation"></a><span data-ttu-id="0745d-113">Реализация</span><span class="sxs-lookup"><span data-stu-id="0745d-113">Implementation</span></span>

<span data-ttu-id="0745d-114">Пример реализации диалогового окна с использованием Dialog API для надстроек Office см. в [этой статье](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) на сайте GitHub.</span><span class="sxs-lookup"><span data-stu-id="0745d-114">For a sample that implements a dialog box, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) in GitHub.</span></span>

## <a name="see-also"></a><span data-ttu-id="0745d-115">См. также</span><span class="sxs-lookup"><span data-stu-id="0745d-115">See also</span></span>

- [<span data-ttu-id="0745d-116">Объект Dialog</span><span class="sxs-lookup"><span data-stu-id="0745d-116">Dialog object</span></span>](https://docs.microsoft.com/javascript/api/office/office.dialog)
- [<span data-ttu-id="0745d-117">Конструктивные шаблоны для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="0745d-117">UX design patterns for Office Add-ins</span></span>](../design/ux-design-pattern-templates.md)


