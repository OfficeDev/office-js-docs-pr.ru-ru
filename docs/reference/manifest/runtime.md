---
title: Среда выполнения в файле манифеста
description: Элемент среды выполнения настраивает надстройку для использования общей среды выполнения JavaScript для различных компонентов, например ленты, области задач, настраиваемых функций.
ms.date: 05/18/2020
localization_priority: Normal
ms.openlocfilehash: c2c404bcaad6e24af58f5c0ed8835343abb97e5f
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278415"
---
# <a name="runtime-element-preview"></a><span data-ttu-id="11cd8-103">Элемент среды выполнения (Предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="11cd8-103">Runtime element (preview)</span></span>

<span data-ttu-id="11cd8-104">Настраивает надстройку для использования общей среды выполнения JavaScript, чтобы различные компоненты запускались в одной среде выполнения.</span><span class="sxs-lookup"><span data-stu-id="11cd8-104">Configures your add-in to use a shared JavaScript runtime so that various components all run in the same runtime.</span></span> <span data-ttu-id="11cd8-105">Дочерний [`<Runtimes>`](runtimes.md) элемент.</span><span class="sxs-lookup"><span data-stu-id="11cd8-105">Child of the [`<Runtimes>`](runtimes.md) element.</span></span>

<span data-ttu-id="11cd8-106">В Excel этот элемент позволяет использовать одну и ту же среду выполнения для ленты, области задач и пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="11cd8-106">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="11cd8-107">Дополнительные сведения можно найти в статье [Настройка надстройки Excel для использования общей среды выполнения JavaScript](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="11cd8-107">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="11cd8-108">В Outlook этот элемент включает активацию надстройки на основе событий.</span><span class="sxs-lookup"><span data-stu-id="11cd8-108">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="11cd8-109">Дополнительные сведения см. [в разделе Настройка надстройки Outlook для активации на основе событий](../../outlook/autolaunch.md).</span><span class="sxs-lookup"><span data-stu-id="11cd8-109">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="11cd8-110">**Тип надстройки:** Область задач, почта</span><span class="sxs-lookup"><span data-stu-id="11cd8-110">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="11cd8-111">**Excel**: общая среда выполнения в настоящее время находится в режиме предварительной версии и доступна только в Excel для Windows.</span><span class="sxs-lookup"><span data-stu-id="11cd8-111">**Excel**: Shared runtime is currently in preview and only available in Excel on Windows.</span></span> <span data-ttu-id="11cd8-112">Для ознакомления с предварительными возможностями необходимо присоединиться к [программе предварительной оценки Office](https://insider.office.com/).</span><span class="sxs-lookup"><span data-stu-id="11cd8-112">To try the preview features, you will need to join [Office Insider](https://insider.office.com/).</span></span>
>
> <span data-ttu-id="11cd8-113">**Outlook**: Активация на основе событий в настоящее время находится [в предварительной версии](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) и доступна только в Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="11cd8-113">**Outlook**: Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="11cd8-114">Дополнительные сведения см. [в статье Просмотр функции активации на основе событий](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span><span class="sxs-lookup"><span data-stu-id="11cd8-114">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="11cd8-115">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="11cd8-115">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="11cd8-116">Содержится в</span><span class="sxs-lookup"><span data-stu-id="11cd8-116">Contained in</span></span>

- [<span data-ttu-id="11cd8-117">Runtimes</span><span class="sxs-lookup"><span data-stu-id="11cd8-117">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="11cd8-118">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="11cd8-118">Attributes</span></span>

|  <span data-ttu-id="11cd8-119">Атрибут</span><span class="sxs-lookup"><span data-stu-id="11cd8-119">Attribute</span></span>  |  <span data-ttu-id="11cd8-120">Обязательный</span><span class="sxs-lookup"><span data-stu-id="11cd8-120">Required</span></span>  |  <span data-ttu-id="11cd8-121">Описание</span><span class="sxs-lookup"><span data-stu-id="11cd8-121">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="11cd8-122">**resid**</span><span class="sxs-lookup"><span data-stu-id="11cd8-122">**resid**</span></span>  |  <span data-ttu-id="11cd8-123">Да</span><span class="sxs-lookup"><span data-stu-id="11cd8-123">Yes</span></span>  | <span data-ttu-id="11cd8-124">Указывает URL-адрес HTML-страницы для надстройки.</span><span class="sxs-lookup"><span data-stu-id="11cd8-124">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="11cd8-125">`resid`Должен сопоставляться с `id` атрибутом `Url` элемента в `Resources` элементе.</span><span class="sxs-lookup"><span data-stu-id="11cd8-125">The `resid` must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |
|  <span data-ttu-id="11cd8-126">**время жизни**</span><span class="sxs-lookup"><span data-stu-id="11cd8-126">**lifetime**</span></span>  |  <span data-ttu-id="11cd8-127">Нет</span><span class="sxs-lookup"><span data-stu-id="11cd8-127">No</span></span>  | <span data-ttu-id="11cd8-128">Значение по умолчанию для свойства `lifetime` `short` и не требуется указывать.</span><span class="sxs-lookup"><span data-stu-id="11cd8-128">The default value for `lifetime` is `short` and doesn't need to be specified.</span></span> <span data-ttu-id="11cd8-129">В надстройках Outlook используется только `short` значение.</span><span class="sxs-lookup"><span data-stu-id="11cd8-129">Outlook add-ins use only the `short` value.</span></span> <span data-ttu-id="11cd8-130">Если вы хотите использовать общую среду выполнения в надстройке Excel, явно задайте для нее значение `long` .</span><span class="sxs-lookup"><span data-stu-id="11cd8-130">If you want to use a shared runtime in an Excel add-in, explicitly set the value to `long`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="11cd8-131">См. также</span><span class="sxs-lookup"><span data-stu-id="11cd8-131">See also</span></span>

- [<span data-ttu-id="11cd8-132">Runtimes</span><span class="sxs-lookup"><span data-stu-id="11cd8-132">Runtimes</span></span>](runtimes.md)
