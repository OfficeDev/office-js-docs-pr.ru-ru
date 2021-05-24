---
title: Время запуска в файле манифеста
description: Элемент Runtime настраивает надстройку для использования общего времени запуска JavaScript для различных компонентов, например ленты, области задач, пользовательских функций.
ms.date: 05/19/2021
localization_priority: Normal
ms.openlocfilehash: cd09abe31ff57eac629c6c61c873c5c886f73f9c
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590916"
---
# <a name="runtime-element"></a><span data-ttu-id="a8fae-103">Элемент runtime</span><span class="sxs-lookup"><span data-stu-id="a8fae-103">Runtime element</span></span>

<span data-ttu-id="a8fae-104">Настраивает надстройку для использования общего времени запуска JavaScript, чтобы все компоненты запускались в одном и том же времени.</span><span class="sxs-lookup"><span data-stu-id="a8fae-104">Configures your add-in to use a shared JavaScript runtime so that various components all run in the same runtime.</span></span> <span data-ttu-id="a8fae-105">Ребенок [`<Runtimes>`](runtimes.md) элемента.</span><span class="sxs-lookup"><span data-stu-id="a8fae-105">Child of the [`<Runtimes>`](runtimes.md) element.</span></span>

<span data-ttu-id="a8fae-106">**Тип надстройки:** Области задач, Почта</span><span class="sxs-lookup"><span data-stu-id="a8fae-106">**Add-in type:** Task pane, Mail</span></span>

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a><span data-ttu-id="a8fae-107">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="a8fae-107">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="a8fae-108">Содержится в</span><span class="sxs-lookup"><span data-stu-id="a8fae-108">Contained in</span></span>

- [<span data-ttu-id="a8fae-109">Runtimes</span><span class="sxs-lookup"><span data-stu-id="a8fae-109">Runtimes</span></span>](runtimes.md)

## <a name="child-elements"></a><span data-ttu-id="a8fae-110">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="a8fae-110">Child elements</span></span>

|  <span data-ttu-id="a8fae-111">Элемент</span><span class="sxs-lookup"><span data-stu-id="a8fae-111">Element</span></span> |  <span data-ttu-id="a8fae-112">Обязательный</span><span class="sxs-lookup"><span data-stu-id="a8fae-112">Required</span></span>  |  <span data-ttu-id="a8fae-113">Описание</span><span class="sxs-lookup"><span data-stu-id="a8fae-113">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="a8fae-114">Override</span><span class="sxs-lookup"><span data-stu-id="a8fae-114">Override</span></span>](override.md) | <span data-ttu-id="a8fae-115">Нет</span><span class="sxs-lookup"><span data-stu-id="a8fae-115">No</span></span> | <span data-ttu-id="a8fae-116">**Outlook:** указывает расположение URL-адреса файла JavaScript, который Outlook для обработчиков точеки [расширения LaunchEvent.](../../reference/manifest/extensionpoint.md#launchevent)</span><span class="sxs-lookup"><span data-stu-id="a8fae-116">**Outlook**: Specifies the URL location of the JavaScript file that Outlook Desktop requires for [LaunchEvent extension point](../../reference/manifest/extensionpoint.md#launchevent) handlers.</span></span> <span data-ttu-id="a8fae-117">**Важно:** в настоящее время можно определить только один элемент и `<Override>` он должен быть типа `javascript` .</span><span class="sxs-lookup"><span data-stu-id="a8fae-117">**Important**: At present, you can only define one `<Override>` element and it must be of type `javascript`.</span></span>|

## <a name="attributes"></a><span data-ttu-id="a8fae-118">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="a8fae-118">Attributes</span></span>

|  <span data-ttu-id="a8fae-119">Атрибут</span><span class="sxs-lookup"><span data-stu-id="a8fae-119">Attribute</span></span>  |  <span data-ttu-id="a8fae-120">Обязательный</span><span class="sxs-lookup"><span data-stu-id="a8fae-120">Required</span></span>  |  <span data-ttu-id="a8fae-121">Описание</span><span class="sxs-lookup"><span data-stu-id="a8fae-121">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="a8fae-122">**resid**</span><span class="sxs-lookup"><span data-stu-id="a8fae-122">**resid**</span></span>  |  <span data-ttu-id="a8fae-123">Да</span><span class="sxs-lookup"><span data-stu-id="a8fae-123">Yes</span></span>  | <span data-ttu-id="a8fae-124">Указывает расположение URL-адреса страницы HTML для надстройки.</span><span class="sxs-lookup"><span data-stu-id="a8fae-124">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="a8fae-125">Символ может быть не более 32 символов и должен соответствовать `resid` `id` атрибуту `Url` элемента `Resources` элемента.</span><span class="sxs-lookup"><span data-stu-id="a8fae-125">The `resid` can be no more than 32 characters and must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |
|  <span data-ttu-id="a8fae-126">**срок службы**</span><span class="sxs-lookup"><span data-stu-id="a8fae-126">**lifetime**</span></span>  |  <span data-ttu-id="a8fae-127">Нет</span><span class="sxs-lookup"><span data-stu-id="a8fae-127">No</span></span>  | <span data-ttu-id="a8fae-128">Значение по умолчанию является и не нужно `lifetime` `short` задано.</span><span class="sxs-lookup"><span data-stu-id="a8fae-128">The default value for `lifetime` is `short` and doesn't need to be specified.</span></span> <span data-ttu-id="a8fae-129">Outlook надстройки используют только `short` значение.</span><span class="sxs-lookup"><span data-stu-id="a8fae-129">Outlook add-ins use only the `short` value.</span></span> <span data-ttu-id="a8fae-130">Если вы хотите использовать совместное время работы в Excel надстройки, явно установите значение `long` .</span><span class="sxs-lookup"><span data-stu-id="a8fae-130">If you want to use a shared runtime in an Excel add-in, explicitly set the value to `long`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="a8fae-131">См. также</span><span class="sxs-lookup"><span data-stu-id="a8fae-131">See also</span></span>

- [<span data-ttu-id="a8fae-132">Runtimes</span><span class="sxs-lookup"><span data-stu-id="a8fae-132">Runtimes</span></span>](runtimes.md)
- [<span data-ttu-id="a8fae-133">Настройка надстройки Office для использования общей среды выполнения JavaScript</span><span class="sxs-lookup"><span data-stu-id="a8fae-133">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="a8fae-134">Настройка надстройки Outlook для активации на основе событий</span><span class="sxs-lookup"><span data-stu-id="a8fae-134">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
