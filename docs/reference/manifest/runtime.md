---
title: Среда выполнения в файле манифеста
description: Элемент среды выполнения настраивает надстройку для использования общей среды выполнения JavaScript для различных компонентов, например ленты, области задач, настраиваемых функций.
ms.date: 05/29/2020
localization_priority: Normal
ms.openlocfilehash: 9e6e13f83db363fb5485c8d8defbc381c80e32d6
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159369"
---
# <a name="runtime-element-preview"></a><span data-ttu-id="aee97-103">Элемент среды выполнения (Предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="aee97-103">Runtime element (preview)</span></span>

<span data-ttu-id="aee97-104">Настраивает надстройку для использования общей среды выполнения JavaScript, чтобы различные компоненты запускались в одной среде выполнения.</span><span class="sxs-lookup"><span data-stu-id="aee97-104">Configures your add-in to use a shared JavaScript runtime so that various components all run in the same runtime.</span></span> <span data-ttu-id="aee97-105">Дочерний [`<Runtimes>`](runtimes.md) элемент.</span><span class="sxs-lookup"><span data-stu-id="aee97-105">Child of the [`<Runtimes>`](runtimes.md) element.</span></span>

<span data-ttu-id="aee97-106">В Excel этот элемент позволяет использовать одну и ту же среду выполнения для ленты, области задач и пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="aee97-106">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="aee97-107">Дополнительные сведения можно найти в статье [Настройка надстройки Excel для использования общей среды выполнения JavaScript](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="aee97-107">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="aee97-108">В Outlook этот элемент включает активацию надстройки на основе событий.</span><span class="sxs-lookup"><span data-stu-id="aee97-108">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="aee97-109">Дополнительные сведения см. [в разделе Настройка надстройки Outlook для активации на основе событий](../../outlook/autolaunch.md).</span><span class="sxs-lookup"><span data-stu-id="aee97-109">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="aee97-110">**Тип надстройки:** Область задач, почта</span><span class="sxs-lookup"><span data-stu-id="aee97-110">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="aee97-111">**Outlook**: Активация на основе событий в настоящее время находится [в предварительной версии](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) и доступна только в Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="aee97-111">**Outlook**: Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="aee97-112">Дополнительные сведения см. [в статье Просмотр функции активации на основе событий](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span><span class="sxs-lookup"><span data-stu-id="aee97-112">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="aee97-113">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="aee97-113">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="aee97-114">Содержится в</span><span class="sxs-lookup"><span data-stu-id="aee97-114">Contained in</span></span>

- [<span data-ttu-id="aee97-115">Runtimes</span><span class="sxs-lookup"><span data-stu-id="aee97-115">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="aee97-116">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="aee97-116">Attributes</span></span>

|  <span data-ttu-id="aee97-117">Атрибут</span><span class="sxs-lookup"><span data-stu-id="aee97-117">Attribute</span></span>  |  <span data-ttu-id="aee97-118">Обязательный</span><span class="sxs-lookup"><span data-stu-id="aee97-118">Required</span></span>  |  <span data-ttu-id="aee97-119">Описание</span><span class="sxs-lookup"><span data-stu-id="aee97-119">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="aee97-120">**resid**</span><span class="sxs-lookup"><span data-stu-id="aee97-120">**resid**</span></span>  |  <span data-ttu-id="aee97-121">Да</span><span class="sxs-lookup"><span data-stu-id="aee97-121">Yes</span></span>  | <span data-ttu-id="aee97-122">Указывает URL-адрес HTML-страницы для надстройки.</span><span class="sxs-lookup"><span data-stu-id="aee97-122">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="aee97-123">`resid`Должен сопоставляться с `id` атрибутом `Url` элемента в `Resources` элементе.</span><span class="sxs-lookup"><span data-stu-id="aee97-123">The `resid` must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |
|  <span data-ttu-id="aee97-124">**время жизни**</span><span class="sxs-lookup"><span data-stu-id="aee97-124">**lifetime**</span></span>  |  <span data-ttu-id="aee97-125">Нет</span><span class="sxs-lookup"><span data-stu-id="aee97-125">No</span></span>  | <span data-ttu-id="aee97-126">Значение по умолчанию для свойства `lifetime` `short` и не требуется указывать.</span><span class="sxs-lookup"><span data-stu-id="aee97-126">The default value for `lifetime` is `short` and doesn't need to be specified.</span></span> <span data-ttu-id="aee97-127">В надстройках Outlook используется только `short` значение.</span><span class="sxs-lookup"><span data-stu-id="aee97-127">Outlook add-ins use only the `short` value.</span></span> <span data-ttu-id="aee97-128">Если вы хотите использовать общую среду выполнения в надстройке Excel, явно задайте для нее значение `long` .</span><span class="sxs-lookup"><span data-stu-id="aee97-128">If you want to use a shared runtime in an Excel add-in, explicitly set the value to `long`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="aee97-129">См. также</span><span class="sxs-lookup"><span data-stu-id="aee97-129">See also</span></span>

- [<span data-ttu-id="aee97-130">Runtimes</span><span class="sxs-lookup"><span data-stu-id="aee97-130">Runtimes</span></span>](runtimes.md)
