---
title: Время запуска в файле манифеста
description: Элемент Runtime настраивает надстройку для использования общего времени запуска JavaScript для различных компонентов, например ленты, области задач, пользовательских функций.
ms.date: 04/08/2021
localization_priority: Normal
ms.openlocfilehash: fa95608d7eff57d68b96ef5b04ec9d33ee63f173
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652246"
---
# <a name="runtime-element"></a><span data-ttu-id="f3dfb-103">Элемент runtime</span><span class="sxs-lookup"><span data-stu-id="f3dfb-103">Runtime element</span></span>

<span data-ttu-id="f3dfb-104">Настраивает надстройку для использования общего времени запуска JavaScript, чтобы все компоненты запускались в одном и том же времени.</span><span class="sxs-lookup"><span data-stu-id="f3dfb-104">Configures your add-in to use a shared JavaScript runtime so that various components all run in the same runtime.</span></span> <span data-ttu-id="f3dfb-105">Ребенок [`<Runtimes>`](runtimes.md) элемента.</span><span class="sxs-lookup"><span data-stu-id="f3dfb-105">Child of the [`<Runtimes>`](runtimes.md) element.</span></span>

<span data-ttu-id="f3dfb-106">**Тип надстройки:** Области задач, Почта</span><span class="sxs-lookup"><span data-stu-id="f3dfb-106">**Add-in type:** Task pane, Mail</span></span>

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a><span data-ttu-id="f3dfb-107">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="f3dfb-107">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="f3dfb-108">Содержится в</span><span class="sxs-lookup"><span data-stu-id="f3dfb-108">Contained in</span></span>

- [<span data-ttu-id="f3dfb-109">Runtimes</span><span class="sxs-lookup"><span data-stu-id="f3dfb-109">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="f3dfb-110">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="f3dfb-110">Attributes</span></span>

|  <span data-ttu-id="f3dfb-111">Атрибут</span><span class="sxs-lookup"><span data-stu-id="f3dfb-111">Attribute</span></span>  |  <span data-ttu-id="f3dfb-112">Обязательный</span><span class="sxs-lookup"><span data-stu-id="f3dfb-112">Required</span></span>  |  <span data-ttu-id="f3dfb-113">Описание</span><span class="sxs-lookup"><span data-stu-id="f3dfb-113">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="f3dfb-114">**resid**</span><span class="sxs-lookup"><span data-stu-id="f3dfb-114">**resid**</span></span>  |  <span data-ttu-id="f3dfb-115">Да</span><span class="sxs-lookup"><span data-stu-id="f3dfb-115">Yes</span></span>  | <span data-ttu-id="f3dfb-116">Указывает расположение URL-адреса страницы HTML для надстройки.</span><span class="sxs-lookup"><span data-stu-id="f3dfb-116">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="f3dfb-117">Символ может быть не более 32 символов и должен соответствовать `resid` `id` атрибуту `Url` элемента `Resources` элемента.</span><span class="sxs-lookup"><span data-stu-id="f3dfb-117">The `resid` can be no more than 32 characters and must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |
|  <span data-ttu-id="f3dfb-118">**срок службы**</span><span class="sxs-lookup"><span data-stu-id="f3dfb-118">**lifetime**</span></span>  |  <span data-ttu-id="f3dfb-119">Нет</span><span class="sxs-lookup"><span data-stu-id="f3dfb-119">No</span></span>  | <span data-ttu-id="f3dfb-120">Значение по умолчанию является и не нужно `lifetime` `short` задано.</span><span class="sxs-lookup"><span data-stu-id="f3dfb-120">The default value for `lifetime` is `short` and doesn't need to be specified.</span></span> <span data-ttu-id="f3dfb-121">Надстройки Outlook используют только `short` значение.</span><span class="sxs-lookup"><span data-stu-id="f3dfb-121">Outlook add-ins use only the `short` value.</span></span> <span data-ttu-id="f3dfb-122">Если вы хотите использовать совместное время работы в надстройки Excel, заочная настройка значения `long` .</span><span class="sxs-lookup"><span data-stu-id="f3dfb-122">If you want to use a shared runtime in an Excel add-in, explicitly set the value to `long`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="f3dfb-123">См. также</span><span class="sxs-lookup"><span data-stu-id="f3dfb-123">See also</span></span>

- [<span data-ttu-id="f3dfb-124">Runtimes</span><span class="sxs-lookup"><span data-stu-id="f3dfb-124">Runtimes</span></span>](runtimes.md)
- [<span data-ttu-id="f3dfb-125">Настройка надстройки Office для использования общей среды выполнения JavaScript</span><span class="sxs-lookup"><span data-stu-id="f3dfb-125">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="f3dfb-126">Настройка надстройки Outlook для активации на основе событий</span><span class="sxs-lookup"><span data-stu-id="f3dfb-126">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
