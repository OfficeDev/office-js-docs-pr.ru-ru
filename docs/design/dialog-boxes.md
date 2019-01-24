---
title: Диалоговые окна в надстройках Office
description: ''
ms.date: 12/04/2017
localization_priority: Priority
ms.openlocfilehash: 78a3419dd93f2a19e3addbeb5a77271b5b124680
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388404"
---
# <a name="dialog-boxes-in-office-add-ins"></a><span data-ttu-id="2c10f-102">Диалоговые окна в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="2c10f-102">Dialog boxes in Office Add-ins</span></span>
 
<span data-ttu-id="2c10f-p101">Диалоговые окна — окна, которые накладываются на активное окно приложения Office. Вы можете использовать диалоговые окна, чтобы показывать страницы входа, которые нельзя открыть непосредственно в области задач, запросы на подтверждение действий, предпринятых пользователем, или видео, которые будут слишком маленькими в области задач.</span><span class="sxs-lookup"><span data-stu-id="2c10f-p101">Dialog boxes are surfaces that float above the active Office application window. You can use dialog boxes to provide additional screen space for tasks such as sign-in pages that can't be opened directly in a task pane or requests to confirm an action taken by a user, or to show videos that might be too small if confined to a task pane.</span></span>

<span data-ttu-id="2c10f-105">*Рисунок 1. Типичный макет диалогового окна*</span><span class="sxs-lookup"><span data-stu-id="2c10f-105">*Figure 1. Typical layout for a dialog box*</span></span>

![Изображение, на котором показан типичный макет диалогового окна](../images/overview-with-app-dialog.png)

## <a name="best-practices"></a><span data-ttu-id="2c10f-107">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="2c10f-107">Best practices</span></span>

|<span data-ttu-id="2c10f-108">**Рекомендуется**</span><span class="sxs-lookup"><span data-stu-id="2c10f-108">**Do**</span></span>|<span data-ttu-id="2c10f-109">**Не рекомендуется**</span><span class="sxs-lookup"><span data-stu-id="2c10f-109">**Don't**</span></span>|
|:-----|:--------|
|<ul><li><span data-ttu-id="2c10f-110">Укажите описательное название, содержащее имя надстройки и название текущей задачи.</span><span class="sxs-lookup"><span data-stu-id="2c10f-110">Include a descriptive title that includes your add-in name along with the current task.</span></span></li></ul>|<ul><li><span data-ttu-id="2c10f-111">Не включайте в него название вашей компании.</span><span class="sxs-lookup"><span data-stu-id="2c10f-111">Don't append your company name to the title.</span></span></li></ul>|
||<ul><li><span data-ttu-id="2c10f-112">Не открывайте диалоговое окно, если этого не требует сценарий.</span><span class="sxs-lookup"><span data-stu-id="2c10f-112">Don't open a dialog box unless the scenario requires it.</span></span></li></ul>|

## <a name="implementation"></a><span data-ttu-id="2c10f-113">Реализация</span><span class="sxs-lookup"><span data-stu-id="2c10f-113">Implementation</span></span>

<span data-ttu-id="2c10f-114">Пример реализации диалогового окна с использованием Dialog API для надстроек Office см. в [этой статье](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) на сайте GitHub.</span><span class="sxs-lookup"><span data-stu-id="2c10f-114">For a sample that implements a dialog box, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) in GitHub.</span></span>

## <a name="see-also"></a><span data-ttu-id="2c10f-115">См. также</span><span class="sxs-lookup"><span data-stu-id="2c10f-115">See also</span></span>

- [<span data-ttu-id="2c10f-116">Ресурсы для разработки на сайте GitHub</span><span class="sxs-lookup"><span data-stu-id="2c10f-116">GitHub Development Resources</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [<span data-ttu-id="2c10f-117">Объект Dialog</span><span class="sxs-lookup"><span data-stu-id="2c10f-117">Dialog object</span></span>](https://docs.microsoft.com/javascript/api/office/office.dialog)


