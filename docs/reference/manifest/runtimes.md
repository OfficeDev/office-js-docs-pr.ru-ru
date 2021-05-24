---
title: Время запуска в файле манифеста
description: Элемент Runtimes указывает время работы надстройки.
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: 80336674c6d954bb9e0c6892feb41cb2f03c5859
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555299"
---
# <a name="runtimes-element"></a><span data-ttu-id="a3f59-103">Элемент Runtimes</span><span class="sxs-lookup"><span data-stu-id="a3f59-103">Runtimes element</span></span>

<span data-ttu-id="a3f59-104">Указывает время запуска надстройки.</span><span class="sxs-lookup"><span data-stu-id="a3f59-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="a3f59-105">Ребенок [`<Host>`](host.md) элемента.</span><span class="sxs-lookup"><span data-stu-id="a3f59-105">Child of the [`<Host>`](host.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="a3f59-106">При работе Office на Windows, надстройка с элементом в манифесте не обязательно будет работать в том же элементе управления веб-просмотром, что и `<Runtimes>` в противном случае.</span><span class="sxs-lookup"><span data-stu-id="a3f59-106">When running in Office on Windows, an add-in that has a `<Runtimes>` element in its manifest does not necessarily run in the same webview control as it otherwise would.</span></span> <span data-ttu-id="a3f59-107">Дополнительные сведения о том, как версии Windows и Office, которые обычно используются для управления [веб-просмотром,](../../concepts/browsers-used-by-office-web-add-ins.md)см. в Office надстройки. Если условия, описанные в нем для Microsoft Edge с webView2 (Chromium на основе), будут выполнены, то надстройка использует этот браузер независимо от того, имеет ли он `<Runtimes>` элемент.</span><span class="sxs-lookup"><span data-stu-id="a3f59-107">For more information about how the versions of Windows and Office determine what webview control is normally used, see [Browsers used by Office Add-ins](../../concepts/browsers-used-by-office-web-add-ins.md). If the conditions described there for using Microsoft Edge with WebView2 (Chromium-based) are met, then the add-in uses that browser whether or not it has a `<Runtimes>` element.</span></span> <span data-ttu-id="a3f59-108">Однако, если эти условия не выполнены, надстройка с элементом всегда использует Internet Explorer 11 независимо от Windows или `<Runtimes>` Microsoft 365 версии.</span><span class="sxs-lookup"><span data-stu-id="a3f59-108">However, when those conditions are not met, an add-in with a `<Runtimes>` element always uses Internet Explorer 11 regardless of the Windows or Microsoft 365 version.</span></span>

<span data-ttu-id="a3f59-109">**Тип надстройки:** Области задач, Почта</span><span class="sxs-lookup"><span data-stu-id="a3f59-109">**Add-in type:** Task pane, Mail</span></span>

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a><span data-ttu-id="a3f59-110">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="a3f59-110">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="a3f59-111">Содержится в</span><span class="sxs-lookup"><span data-stu-id="a3f59-111">Contained in</span></span>

[<span data-ttu-id="a3f59-112">Host</span><span class="sxs-lookup"><span data-stu-id="a3f59-112">Host</span></span>](host.md)

## <a name="child-elements"></a><span data-ttu-id="a3f59-113">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="a3f59-113">Child elements</span></span>

|  <span data-ttu-id="a3f59-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="a3f59-114">Element</span></span> |  <span data-ttu-id="a3f59-115">Обязательный</span><span class="sxs-lookup"><span data-stu-id="a3f59-115">Required</span></span>  |  <span data-ttu-id="a3f59-116">Описание</span><span class="sxs-lookup"><span data-stu-id="a3f59-116">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="a3f59-117">Время выполнения</span><span class="sxs-lookup"><span data-stu-id="a3f59-117">Runtime</span></span>](runtime.md) | <span data-ttu-id="a3f59-118">Да</span><span class="sxs-lookup"><span data-stu-id="a3f59-118">Yes</span></span> |  <span data-ttu-id="a3f59-119">Время запуска надстройки.</span><span class="sxs-lookup"><span data-stu-id="a3f59-119">The runtime for your add-in.</span></span> <span data-ttu-id="a3f59-120">**Важно.** В настоящее время можно определить только один `<Runtime>` элемент.</span><span class="sxs-lookup"><span data-stu-id="a3f59-120">**Important**: At present, you can only define one `<Runtime>` element.</span></span> |

## <a name="see-also"></a><span data-ttu-id="a3f59-121">См. также</span><span class="sxs-lookup"><span data-stu-id="a3f59-121">See also</span></span>

- [<span data-ttu-id="a3f59-122">Время выполнения</span><span class="sxs-lookup"><span data-stu-id="a3f59-122">Runtime</span></span>](runtime.md)
- [<span data-ttu-id="a3f59-123">Настройка надстройки Office для использования общей среды выполнения JavaScript</span><span class="sxs-lookup"><span data-stu-id="a3f59-123">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="a3f59-124">Настройка надстройки Outlook для активации на основе событий</span><span class="sxs-lookup"><span data-stu-id="a3f59-124">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
