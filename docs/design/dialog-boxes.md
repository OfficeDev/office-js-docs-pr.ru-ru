---
title: Диалоговые окна в надстройках Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: f18f603d76a902bdce56152ecb3f63bbafad56fb
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2018
ms.locfileid: "23945752"
---
# <a name="dialog-boxes-in-office-add-ins"></a><span data-ttu-id="dc7a2-102">Диалоговые окна в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="dc7a2-102">Dialog boxes in Office Add-ins</span></span>
 
<span data-ttu-id="dc7a2-p101">Диалоговые окна — окна, которые накладываются на активное окно приложения Office. Вы можете использовать диалоговые окна, чтобы показывать страницы входа, которые нельзя открыть непосредственно в области задач, запросы на подтверждение действий, предпринятых пользователем, или видео, которые будут слишком маленькими в области задач.</span><span class="sxs-lookup"><span data-stu-id="dc7a2-p101">Dialog boxes are surfaces that float above the active Office application window. You can use dialog boxes to provide additional screen space for tasks such as sign-in pages that can't be opened directly in a task pane or requests to confirm an action taken by a user, or to show videos that might be too small if confined to a task pane.</span></span>

<span data-ttu-id="dc7a2-105">*Рисунок 1. Типичный макет диалогового окна*</span><span class="sxs-lookup"><span data-stu-id="dc7a2-105">*Figure 1. Typical layout for a dialog box*</span></span>

![Изображение, на котором показан типичный макет диалогового окна](../images/overview-with-app-dialog.png)

## <a name="best-practices"></a><span data-ttu-id="dc7a2-107">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="dc7a2-107">Best practices</span></span>

|<span data-ttu-id="dc7a2-108">**Рекомендуется**</span><span class="sxs-lookup"><span data-stu-id="dc7a2-108">**Do**</span></span>|<span data-ttu-id="dc7a2-109">**Не рекомендуется**</span><span class="sxs-lookup"><span data-stu-id="dc7a2-109">**Don't**</span></span>|
|:-----|:--------|
|<ul><li><span data-ttu-id="dc7a2-110">Укажите описательное название, содержащее имя надстройки и название текущей задачи.</span><span class="sxs-lookup"><span data-stu-id="dc7a2-110">Include a descriptive title that includes your add-in name along with the current task.</span></span></li></ul>|<ul><li><span data-ttu-id="dc7a2-111">Не включайте в него название вашей компании.</span><span class="sxs-lookup"><span data-stu-id="dc7a2-111">Don't append your company name to the title.</span></span></li></ul>|
||<ul><li><span data-ttu-id="dc7a2-112">Не открывайте диалоговое окно, если этого не требует сценарий.</span><span class="sxs-lookup"><span data-stu-id="dc7a2-112">Don't open a dialog box unless the scenario requires it.</span></span></li></ul>|

## <a name="implementation"></a><span data-ttu-id="dc7a2-113">Реализация</span><span class="sxs-lookup"><span data-stu-id="dc7a2-113">Implementation</span></span>

<span data-ttu-id="dc7a2-114">Пример реализации диалогового окна с использованием Dialog API для надстроек Office см. в [этой статье](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) на сайте GitHub.</span><span class="sxs-lookup"><span data-stu-id="dc7a2-114">For a sample that implements a dialog box, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example) in GitHub.</span></span>

## <a name="see-also"></a><span data-ttu-id="dc7a2-115">См. также</span><span class="sxs-lookup"><span data-stu-id="dc7a2-115">See also</span></span>

- [<span data-ttu-id="dc7a2-116">Пример шаблона для взаимодействия с пользователем</span><span class="sxs-lookup"><span data-stu-id="dc7a2-116">UX Pattern Sample</span></span>](https://office.visualstudio.com/DefaultCollection/OC/_git/GettingStarted-FabricReact)
- [<span data-ttu-id="dc7a2-117">Ресурсы для разработки на сайте GitHub</span><span class="sxs-lookup"><span data-stu-id="dc7a2-117">GitHub Development Resources</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [<span data-ttu-id="dc7a2-118">Объект Dialog</span><span class="sxs-lookup"><span data-stu-id="dc7a2-118">Dialog object</span></span>](https://docs.microsoft.com/javascript/api/office/office.dialog?view=office-js)


