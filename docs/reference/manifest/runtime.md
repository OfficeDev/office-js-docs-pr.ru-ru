---
title: Время выполнения в файле манифеста
description: Элемент Runtime настраивает надстройки для использования общего времени выполнения JavaScript для различных компонентов, например ленты, панели задач, пользовательских функций.
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: c59e5a23e53940aea46c758d710b4a455cb5c0cc
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555306"
---
# <a name="runtime-element"></a><span data-ttu-id="05a1a-103">Элемент времени выполнения</span><span class="sxs-lookup"><span data-stu-id="05a1a-103">Runtime element</span></span>

<span data-ttu-id="05a1a-104">Настраивает надстройки для использования общего времени выполнения JavaScript, чтобы все различные компоненты запускаются в одно и то же время выполнения.</span><span class="sxs-lookup"><span data-stu-id="05a1a-104">Configures your add-in to use a shared JavaScript runtime so that various components all run in the same runtime.</span></span> <span data-ttu-id="05a1a-105">Дитя [`<Runtimes>`](runtimes.md) элемента.</span><span class="sxs-lookup"><span data-stu-id="05a1a-105">Child of the [`<Runtimes>`](runtimes.md) element.</span></span>

<span data-ttu-id="05a1a-106">**Тип дополнения:** Панель задач, Почта</span><span class="sxs-lookup"><span data-stu-id="05a1a-106">**Add-in type:** Task pane, Mail</span></span>

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a><span data-ttu-id="05a1a-107">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="05a1a-107">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="05a1a-108">Содержится в</span><span class="sxs-lookup"><span data-stu-id="05a1a-108">Contained in</span></span>

- [<span data-ttu-id="05a1a-109">Runtimes</span><span class="sxs-lookup"><span data-stu-id="05a1a-109">Runtimes</span></span>](runtimes.md)

## <a name="child-elements"></a><span data-ttu-id="05a1a-110">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="05a1a-110">Child elements</span></span>

|  <span data-ttu-id="05a1a-111">Элемент</span><span class="sxs-lookup"><span data-stu-id="05a1a-111">Element</span></span> |  <span data-ttu-id="05a1a-112">Обязательный</span><span class="sxs-lookup"><span data-stu-id="05a1a-112">Required</span></span>  |  <span data-ttu-id="05a1a-113">Описание</span><span class="sxs-lookup"><span data-stu-id="05a1a-113">Description</span></span>  |
|:-----|:-----|:-----|
| <span data-ttu-id="05a1a-114">[Переопределение](override.md) (предварительный просмотр)</span><span class="sxs-lookup"><span data-stu-id="05a1a-114">[Override](override.md) (preview)</span></span> | <span data-ttu-id="05a1a-115">Нет</span><span class="sxs-lookup"><span data-stu-id="05a1a-115">No</span></span> | <span data-ttu-id="05a1a-116">**Outlook**: Определяет местоположение URL-адреса файла JavaScript, который требуется Outlook Desktop [для обработчиков токов точки расширения LaunchEvent.](../../reference/manifest/extensionpoint.md#launchevent-preview)</span><span class="sxs-lookup"><span data-stu-id="05a1a-116">**Outlook**: Specifies the URL location of the JavaScript file that Outlook Desktop requires for [LaunchEvent extension point](../../reference/manifest/extensionpoint.md#launchevent-preview) handlers.</span></span> <span data-ttu-id="05a1a-117">**Важно:** В настоящее время вы можете определить только `<Override>` один элемент, и он должен быть типа `javascript` .</span><span class="sxs-lookup"><span data-stu-id="05a1a-117">**Important**: At present, you can only define one `<Override>` element and it must be of type `javascript`.</span></span>|

## <a name="attributes"></a><span data-ttu-id="05a1a-118">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="05a1a-118">Attributes</span></span>

|  <span data-ttu-id="05a1a-119">Атрибут</span><span class="sxs-lookup"><span data-stu-id="05a1a-119">Attribute</span></span>  |  <span data-ttu-id="05a1a-120">Обязательный</span><span class="sxs-lookup"><span data-stu-id="05a1a-120">Required</span></span>  |  <span data-ttu-id="05a1a-121">Описание</span><span class="sxs-lookup"><span data-stu-id="05a1a-121">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="05a1a-122">**resid**</span><span class="sxs-lookup"><span data-stu-id="05a1a-122">**resid**</span></span>  |  <span data-ttu-id="05a1a-123">Да</span><span class="sxs-lookup"><span data-stu-id="05a1a-123">Yes</span></span>  | <span data-ttu-id="05a1a-124">Определяет местоположение URL страницы HTML для надстройки.</span><span class="sxs-lookup"><span data-stu-id="05a1a-124">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="05a1a-125">Может `resid` быть не более 32 символов и должен `id` соответствовать атрибуту `Url` элемента в `Resources` элементе.</span><span class="sxs-lookup"><span data-stu-id="05a1a-125">The `resid` can be no more than 32 characters and must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |
|  <span data-ttu-id="05a1a-126">**продолжительность жизни**</span><span class="sxs-lookup"><span data-stu-id="05a1a-126">**lifetime**</span></span>  |  <span data-ttu-id="05a1a-127">Нет</span><span class="sxs-lookup"><span data-stu-id="05a1a-127">No</span></span>  | <span data-ttu-id="05a1a-128">Значение по `lifetime` умолчанию `short` для является и не должно быть указано.</span><span class="sxs-lookup"><span data-stu-id="05a1a-128">The default value for `lifetime` is `short` and doesn't need to be specified.</span></span> <span data-ttu-id="05a1a-129">Outlook надстройки используют только `short` значение.</span><span class="sxs-lookup"><span data-stu-id="05a1a-129">Outlook add-ins use only the `short` value.</span></span> <span data-ttu-id="05a1a-130">Если вы хотите использовать общее время выполнения в Excel в дополнение, явно установите `long` значение.</span><span class="sxs-lookup"><span data-stu-id="05a1a-130">If you want to use a shared runtime in an Excel add-in, explicitly set the value to `long`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="05a1a-131">См. также</span><span class="sxs-lookup"><span data-stu-id="05a1a-131">See also</span></span>

- [<span data-ttu-id="05a1a-132">Runtimes</span><span class="sxs-lookup"><span data-stu-id="05a1a-132">Runtimes</span></span>](runtimes.md)
- [<span data-ttu-id="05a1a-133">Настройка надстройки Office для использования общей среды выполнения JavaScript</span><span class="sxs-lookup"><span data-stu-id="05a1a-133">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="05a1a-134">Настройте Outlook для активации на основе событий</span><span class="sxs-lookup"><span data-stu-id="05a1a-134">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
