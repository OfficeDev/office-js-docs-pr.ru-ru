---
title: Время запуска в файле манифеста
description: Элемент Runtimes указывает время работы надстройки.
ms.date: 04/16/2021
localization_priority: Normal
ms.openlocfilehash: 8f4a602c05b9af7bde9f644ef40b61a214e66cd5
ms.sourcegitcommit: da8ad214406f2e1cd80982af8a13090e76187dbd
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/21/2021
ms.locfileid: "51917088"
---
# <a name="runtimes-element"></a><span data-ttu-id="0ccd0-103">Элемент Runtimes</span><span class="sxs-lookup"><span data-stu-id="0ccd0-103">Runtimes element</span></span>

<span data-ttu-id="0ccd0-104">Указывает время запуска надстройки.</span><span class="sxs-lookup"><span data-stu-id="0ccd0-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="0ccd0-105">Ребенок [`<Host>`](host.md) элемента.</span><span class="sxs-lookup"><span data-stu-id="0ccd0-105">Child of the [`<Host>`](host.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="0ccd0-106">При работе в Office on Windows надстройка с элементом манифеста не обязательно будет работать в том же элементе управления веб-просмотром, что и `<Runtimes>` в противном случае.</span><span class="sxs-lookup"><span data-stu-id="0ccd0-106">When running in Office on Windows, an add-in that has a `<Runtimes>` element in its manifest does not necessarily run in the same webview control as it otherwise would.</span></span> <span data-ttu-id="0ccd0-107">Дополнительные сведения о том, как версии Windows и Office определяют, как обычно используется управление веб-просмотром, см. в браузерах, используемых [надстройки Office.](../../concepts/browsers-used-by-office-web-add-ins.md) Если условия, описанные там для использования Microsoft Edge с WebView2 (на основе хрома), выполнены, то надстройка использует этот браузер независимо от того, имеет ли он `<Runtimes>` элемент.</span><span class="sxs-lookup"><span data-stu-id="0ccd0-107">For more information about how the versions of Windows and Office determine what webview control is normally used, see [Browsers used by Office Add-ins](../../concepts/browsers-used-by-office-web-add-ins.md). If the conditions described there for using Microsoft Edge with WebView2 (Chromium-based) are met, then the add-in uses that browser whether or not it has a `<Runtimes>` element.</span></span> <span data-ttu-id="0ccd0-108">Однако, если эти условия не выполнены, надстройка с элементом всегда использует Internet Explorer 11 независимо от версии Windows или `<Runtimes>` Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="0ccd0-108">However, when those conditions are not met, an add-in with a `<Runtimes>` element always uses Internet Explorer 11 regardless of the Windows or Microsoft 365 version.</span></span>

<span data-ttu-id="0ccd0-109">**Тип надстройки:** Области задач, Почта</span><span class="sxs-lookup"><span data-stu-id="0ccd0-109">**Add-in type:** Task pane, Mail</span></span>

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a><span data-ttu-id="0ccd0-110">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="0ccd0-110">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="0ccd0-111">Содержится в</span><span class="sxs-lookup"><span data-stu-id="0ccd0-111">Contained in</span></span>

[<span data-ttu-id="0ccd0-112">Host</span><span class="sxs-lookup"><span data-stu-id="0ccd0-112">Host</span></span>](host.md)

## <a name="child-elements"></a><span data-ttu-id="0ccd0-113">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="0ccd0-113">Child elements</span></span>

|  <span data-ttu-id="0ccd0-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="0ccd0-114">Element</span></span> |  <span data-ttu-id="0ccd0-115">Обязательный</span><span class="sxs-lookup"><span data-stu-id="0ccd0-115">Required</span></span>  |  <span data-ttu-id="0ccd0-116">Описание</span><span class="sxs-lookup"><span data-stu-id="0ccd0-116">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="0ccd0-117">Время выполнения</span><span class="sxs-lookup"><span data-stu-id="0ccd0-117">Runtime</span></span>](runtime.md) | <span data-ttu-id="0ccd0-118">Да</span><span class="sxs-lookup"><span data-stu-id="0ccd0-118">Yes</span></span> |  <span data-ttu-id="0ccd0-119">Время запуска надстройки.</span><span class="sxs-lookup"><span data-stu-id="0ccd0-119">The runtime for your add-in.</span></span> |

## <a name="see-also"></a><span data-ttu-id="0ccd0-120">См. также</span><span class="sxs-lookup"><span data-stu-id="0ccd0-120">See also</span></span>

- [<span data-ttu-id="0ccd0-121">Время выполнения</span><span class="sxs-lookup"><span data-stu-id="0ccd0-121">Runtime</span></span>](runtime.md)
- [<span data-ttu-id="0ccd0-122">Настройка надстройки Office для использования общей среды выполнения JavaScript</span><span class="sxs-lookup"><span data-stu-id="0ccd0-122">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="0ccd0-123">Настройка надстройки Outlook для активации на основе событий</span><span class="sxs-lookup"><span data-stu-id="0ccd0-123">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
