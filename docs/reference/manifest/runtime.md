---
title: Runtime в файле манифеста
description: Элемент runtime настраивает надстройку на использование общей компоненты javaScript для различных компонентов, например ленты, области задач, пользовательских функций.
ms.date: 05/29/2020
localization_priority: Normal
ms.openlocfilehash: 3cabfacc665ccf6c0e4e796cb0e1fbc70c770ee3
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789186"
---
# <a name="runtime-element-preview"></a><span data-ttu-id="183db-103">Элемент runtime (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="183db-103">Runtime element (preview)</span></span>

<span data-ttu-id="183db-104">Настраивает надстройку для использования общей времени работы JavaScript, чтобы все компоненты запускались в одной среде.</span><span class="sxs-lookup"><span data-stu-id="183db-104">Configures your add-in to use a shared JavaScript runtime so that various components all run in the same runtime.</span></span> <span data-ttu-id="183db-105">Child of the [`<Runtimes>`](runtimes.md) element.</span><span class="sxs-lookup"><span data-stu-id="183db-105">Child of the [`<Runtimes>`](runtimes.md) element.</span></span>

<span data-ttu-id="183db-106">В Excel этот элемент позволяет ленте, области задач и пользовательским функциям использовать ту же времени работы.</span><span class="sxs-lookup"><span data-stu-id="183db-106">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="183db-107">Дополнительные сведения см. в настройках надстройки Excel для использования общей времени [работы JavaScript.](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)</span><span class="sxs-lookup"><span data-stu-id="183db-107">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="183db-108">В Outlook этот элемент включает активацию надстройки на основе событий.</span><span class="sxs-lookup"><span data-stu-id="183db-108">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="183db-109">Дополнительные сведения см. в настройке [надстройки Outlook для активации на основе событий.](../../outlook/autolaunch.md)</span><span class="sxs-lookup"><span data-stu-id="183db-109">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="183db-110">**Тип надстройки:** Области задач, почта</span><span class="sxs-lookup"><span data-stu-id="183db-110">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="183db-111">**Outlook**: активация на основе событий в настоящее время находится в [предварительной](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) версии и доступна только в Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="183db-111">**Outlook**: Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="183db-112">Дополнительные сведения см. в [предварительном просмотре функции активации на основе событий.](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)</span><span class="sxs-lookup"><span data-stu-id="183db-112">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="183db-113">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="183db-113">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="183db-114">Содержится в</span><span class="sxs-lookup"><span data-stu-id="183db-114">Contained in</span></span>

- [<span data-ttu-id="183db-115">Runtimes</span><span class="sxs-lookup"><span data-stu-id="183db-115">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="183db-116">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="183db-116">Attributes</span></span>

|  <span data-ttu-id="183db-117">Атрибут</span><span class="sxs-lookup"><span data-stu-id="183db-117">Attribute</span></span>  |  <span data-ttu-id="183db-118">Обязательный</span><span class="sxs-lookup"><span data-stu-id="183db-118">Required</span></span>  |  <span data-ttu-id="183db-119">Описание</span><span class="sxs-lookup"><span data-stu-id="183db-119">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="183db-120">**resid**</span><span class="sxs-lookup"><span data-stu-id="183db-120">**resid**</span></span>  |  <span data-ttu-id="183db-121">Да</span><span class="sxs-lookup"><span data-stu-id="183db-121">Yes</span></span>  | <span data-ttu-id="183db-122">Указывает URL-адрес HTML-страницы для надстройки.</span><span class="sxs-lookup"><span data-stu-id="183db-122">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="183db-123">Он может иметь не более 32 символов и должен соответствовать `resid` `id` атрибуту `Url` элемента в `Resources` элементе.</span><span class="sxs-lookup"><span data-stu-id="183db-123">The `resid` can be no more than 32 characters and must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |
|  <span data-ttu-id="183db-124">**lifetime**</span><span class="sxs-lookup"><span data-stu-id="183db-124">**lifetime**</span></span>  |  <span data-ttu-id="183db-125">Нет</span><span class="sxs-lookup"><span data-stu-id="183db-125">No</span></span>  | <span data-ttu-id="183db-126">Значение по умолчанию : и не `lifetime` `short` требуется быть заданным.</span><span class="sxs-lookup"><span data-stu-id="183db-126">The default value for `lifetime` is `short` and doesn't need to be specified.</span></span> <span data-ttu-id="183db-127">Надстройки Outlook используют только `short` значение.</span><span class="sxs-lookup"><span data-stu-id="183db-127">Outlook add-ins use only the `short` value.</span></span> <span data-ttu-id="183db-128">Если вы хотите использовать общую time runtime в надстройки Excel, явно установите значение `long` .</span><span class="sxs-lookup"><span data-stu-id="183db-128">If you want to use a shared runtime in an Excel add-in, explicitly set the value to `long`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="183db-129">См. также</span><span class="sxs-lookup"><span data-stu-id="183db-129">See also</span></span>

- [<span data-ttu-id="183db-130">Runtimes</span><span class="sxs-lookup"><span data-stu-id="183db-130">Runtimes</span></span>](runtimes.md)
