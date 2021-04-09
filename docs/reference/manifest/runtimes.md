---
title: Время запуска в файле манифеста
description: Элемент Runtimes указывает время работы надстройки.
ms.date: 04/08/2021
localization_priority: Normal
ms.openlocfilehash: a5cd05a0890615375bf3466caf70d22f9912d951
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652238"
---
# <a name="runtimes-element"></a><span data-ttu-id="f7d31-103">Элемент Runtimes</span><span class="sxs-lookup"><span data-stu-id="f7d31-103">Runtimes element</span></span>

<span data-ttu-id="f7d31-104">Указывает время запуска надстройки.</span><span class="sxs-lookup"><span data-stu-id="f7d31-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="f7d31-105">Ребенок [`<Host>`](host.md) элемента.</span><span class="sxs-lookup"><span data-stu-id="f7d31-105">Child of the [`<Host>`](host.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="f7d31-106">При работе в Office на Windows надстройка использует браузер Internet Explorer 11.</span><span class="sxs-lookup"><span data-stu-id="f7d31-106">When running in Office on Windows, your add-in uses the Internet Explorer 11 browser.</span></span>

<span data-ttu-id="f7d31-107">**Тип надстройки:** Области задач, Почта</span><span class="sxs-lookup"><span data-stu-id="f7d31-107">**Add-in type:** Task pane, Mail</span></span>

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a><span data-ttu-id="f7d31-108">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="f7d31-108">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="f7d31-109">Содержится в</span><span class="sxs-lookup"><span data-stu-id="f7d31-109">Contained in</span></span>

[<span data-ttu-id="f7d31-110">Host</span><span class="sxs-lookup"><span data-stu-id="f7d31-110">Host</span></span>](host.md)

## <a name="child-elements"></a><span data-ttu-id="f7d31-111">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="f7d31-111">Child elements</span></span>

|  <span data-ttu-id="f7d31-112">Элемент</span><span class="sxs-lookup"><span data-stu-id="f7d31-112">Element</span></span> |  <span data-ttu-id="f7d31-113">Обязательный</span><span class="sxs-lookup"><span data-stu-id="f7d31-113">Required</span></span>  |  <span data-ttu-id="f7d31-114">Описание</span><span class="sxs-lookup"><span data-stu-id="f7d31-114">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="f7d31-115">Runtime</span><span class="sxs-lookup"><span data-stu-id="f7d31-115">Runtime</span></span>](runtime.md) | <span data-ttu-id="f7d31-116">Да</span><span class="sxs-lookup"><span data-stu-id="f7d31-116">Yes</span></span> |  <span data-ttu-id="f7d31-117">Время запуска надстройки.</span><span class="sxs-lookup"><span data-stu-id="f7d31-117">The runtime for your add-in.</span></span> |

## <a name="see-also"></a><span data-ttu-id="f7d31-118">См. также</span><span class="sxs-lookup"><span data-stu-id="f7d31-118">See also</span></span>

- [<span data-ttu-id="f7d31-119">Runtime</span><span class="sxs-lookup"><span data-stu-id="f7d31-119">Runtime</span></span>](runtime.md)
- [<span data-ttu-id="f7d31-120">Настройка надстройки Office для использования общей среды выполнения JavaScript</span><span class="sxs-lookup"><span data-stu-id="f7d31-120">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="f7d31-121">Настройка надстройки Outlook для активации на основе событий</span><span class="sxs-lookup"><span data-stu-id="f7d31-121">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
