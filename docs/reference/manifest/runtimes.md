---
title: Время времени времени времени времени в файле манифеста
description: Элемент Runtimes определяет время выполнения надстройки.
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: 80336674c6d954bb9e0c6892feb41cb2f03c5859
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555299"
---
# <a name="runtimes-element"></a><span data-ttu-id="7d4dd-103">Элемент времени бегут</span><span class="sxs-lookup"><span data-stu-id="7d4dd-103">Runtimes element</span></span>

<span data-ttu-id="7d4dd-104">Определяет время выполнения надстройки.</span><span class="sxs-lookup"><span data-stu-id="7d4dd-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="7d4dd-105">Дитя [`<Host>`](host.md) элемента.</span><span class="sxs-lookup"><span data-stu-id="7d4dd-105">Child of the [`<Host>`](host.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="7d4dd-106">При запуске Office на Windows, надстройку, которая имеет элемент в манифесте, не обязательно работает в том же `<Runtimes>` контроле веб-вида, как это было бы в противном случае.</span><span class="sxs-lookup"><span data-stu-id="7d4dd-106">When running in Office on Windows, an add-in that has a `<Runtimes>` element in its manifest does not necessarily run in the same webview control as it otherwise would.</span></span> <span data-ttu-id="7d4dd-107">Для получения дополнительной информации о том, как Windows и Office определяют, какой элемент управления веб-видом обычно [используется, Office см.](../../concepts/browsers-used-by-office-web-add-ins.md) Если описанные там условия для использования Microsoft Edge с WebView2 (Chromium-based) выполнены, то надстройок использует этот браузер независимо от того, есть ли у него `<Runtimes>` элемент.</span><span class="sxs-lookup"><span data-stu-id="7d4dd-107">For more information about how the versions of Windows and Office determine what webview control is normally used, see [Browsers used by Office Add-ins](../../concepts/browsers-used-by-office-web-add-ins.md). If the conditions described there for using Microsoft Edge with WebView2 (Chromium-based) are met, then the add-in uses that browser whether or not it has a `<Runtimes>` element.</span></span> <span data-ttu-id="7d4dd-108">Однако, когда эти условия не выполнены, надстройа с `<Runtimes>` элементом всегда использует Internet Explorer 11 независимо от Windows или Microsoft 365 версии.</span><span class="sxs-lookup"><span data-stu-id="7d4dd-108">However, when those conditions are not met, an add-in with a `<Runtimes>` element always uses Internet Explorer 11 regardless of the Windows or Microsoft 365 version.</span></span>

<span data-ttu-id="7d4dd-109">**Тип дополнения:** Панель задач, Почта</span><span class="sxs-lookup"><span data-stu-id="7d4dd-109">**Add-in type:** Task pane, Mail</span></span>

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a><span data-ttu-id="7d4dd-110">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="7d4dd-110">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="7d4dd-111">Содержится в</span><span class="sxs-lookup"><span data-stu-id="7d4dd-111">Contained in</span></span>

[<span data-ttu-id="7d4dd-112">Host</span><span class="sxs-lookup"><span data-stu-id="7d4dd-112">Host</span></span>](host.md)

## <a name="child-elements"></a><span data-ttu-id="7d4dd-113">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="7d4dd-113">Child elements</span></span>

|  <span data-ttu-id="7d4dd-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="7d4dd-114">Element</span></span> |  <span data-ttu-id="7d4dd-115">Обязательный</span><span class="sxs-lookup"><span data-stu-id="7d4dd-115">Required</span></span>  |  <span data-ttu-id="7d4dd-116">Описание</span><span class="sxs-lookup"><span data-stu-id="7d4dd-116">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="7d4dd-117">Время выполнения</span><span class="sxs-lookup"><span data-stu-id="7d4dd-117">Runtime</span></span>](runtime.md) | <span data-ttu-id="7d4dd-118">Да</span><span class="sxs-lookup"><span data-stu-id="7d4dd-118">Yes</span></span> |  <span data-ttu-id="7d4dd-119">Время выполнения надстройок.</span><span class="sxs-lookup"><span data-stu-id="7d4dd-119">The runtime for your add-in.</span></span> <span data-ttu-id="7d4dd-120">**Важно**: В настоящее время можно определить только один `<Runtime>` элемент.</span><span class="sxs-lookup"><span data-stu-id="7d4dd-120">**Important**: At present, you can only define one `<Runtime>` element.</span></span> |

## <a name="see-also"></a><span data-ttu-id="7d4dd-121">См. также</span><span class="sxs-lookup"><span data-stu-id="7d4dd-121">See also</span></span>

- [<span data-ttu-id="7d4dd-122">Время выполнения</span><span class="sxs-lookup"><span data-stu-id="7d4dd-122">Runtime</span></span>](runtime.md)
- [<span data-ttu-id="7d4dd-123">Настройка надстройки Office для использования общей среды выполнения JavaScript</span><span class="sxs-lookup"><span data-stu-id="7d4dd-123">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="7d4dd-124">Настройте Outlook для активации на основе событий</span><span class="sxs-lookup"><span data-stu-id="7d4dd-124">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
