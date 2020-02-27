---
title: Среда выполнения в файле манифеста (Предварительная версия)
description: ''
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: 26702896604f9ecf4c69296e5110efe5cdf4218b
ms.sourcegitcommit: dd6d00202f6466c27418247dad7bd136555a6036
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/26/2020
ms.locfileid: "42283886"
---
# <a name="runtime-element-preview"></a><span data-ttu-id="b20ae-102">Элемент среды выполнения (Предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="b20ae-102">Runtime element (preview)</span></span>

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="b20ae-103">Дочерний элемент [`<Runtimes>`](runtimes.md) элемента.</span><span class="sxs-lookup"><span data-stu-id="b20ae-103">Child element of the [`<Runtimes>`](runtimes.md) element.</span></span> <span data-ttu-id="b20ae-104">Этот элемент настраивает надстройку, чтобы использовать общую среду выполнения JavaScript, чтобы Ваша лента, область задач и пользовательские функции выполнялись в одной среде выполнения.</span><span class="sxs-lookup"><span data-stu-id="b20ae-104">This element configures your add-in to use a shared JavaScript runtime so that your ribbon, task pane, and custom functions, all run in the same runtime.</span></span> <span data-ttu-id="b20ae-105">Дополнительные сведения можно найти в статье [Настройка надстройки Excel для использования общей среды выполнения JavaScript](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="b20ae-105">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="b20ae-106">**Тип надстройки:** надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="b20ae-106">**Add-in type:** Task pane</span></span>

> [!IMPORTANT]
<span data-ttu-id="b20ae-107">В настоящее время общедоступная среда выполнения <<<<<<< для ГОЛОВного общего доступа доступна только в Excel для Windows.</span><span class="sxs-lookup"><span data-stu-id="b20ae-107"><<<<<<< HEAD Shared runtime is currently in preview and are only available on Excel on Windows.</span></span> <span data-ttu-id="b20ae-108">Для ознакомления с предварительными возможностями необходимо присоединиться к [программе предварительной оценки Office](https://insider.office.com/).</span><span class="sxs-lookup"><span data-stu-id="b20ae-108">To try the preview features, you will need to join [Office Insider](https://insider.office.com/).</span></span>

## <a name="syntax"></a><span data-ttu-id="b20ae-109">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="b20ae-109">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="b20ae-110">Содержится в</span><span class="sxs-lookup"><span data-stu-id="b20ae-110">Contained in</span></span>

- [<span data-ttu-id="b20ae-111">Runtimes</span><span class="sxs-lookup"><span data-stu-id="b20ae-111">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="b20ae-112">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b20ae-112">Attributes</span></span>

|  <span data-ttu-id="b20ae-113">Атрибут</span><span class="sxs-lookup"><span data-stu-id="b20ae-113">Attribute</span></span>  |  <span data-ttu-id="b20ae-114">Обязательный</span><span class="sxs-lookup"><span data-stu-id="b20ae-114">Required</span></span>  |  <span data-ttu-id="b20ae-115">Описание</span><span class="sxs-lookup"><span data-stu-id="b20ae-115">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="b20ae-116">**время жизни = "Long"**</span><span class="sxs-lookup"><span data-stu-id="b20ae-116">**lifetime="long"**</span></span>  |  <span data-ttu-id="b20ae-117">Да</span><span class="sxs-lookup"><span data-stu-id="b20ae-117">Yes</span></span>  | <span data-ttu-id="b20ae-118">Всегда следует использовать `long` , если вы хотите использовать общую среду выполнения для надстройки Excel.</span><span class="sxs-lookup"><span data-stu-id="b20ae-118">Should always be `long` if you want to use a shared runtime for the Excel add-in.</span></span> |
|  <span data-ttu-id="b20ae-119">**resid**</span><span class="sxs-lookup"><span data-stu-id="b20ae-119">**resid**</span></span>  |  <span data-ttu-id="b20ae-120">Да</span><span class="sxs-lookup"><span data-stu-id="b20ae-120">Yes</span></span>  | <span data-ttu-id="b20ae-121">Указывает URL-адрес HTML-страницы для надстройки.</span><span class="sxs-lookup"><span data-stu-id="b20ae-121">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="b20ae-122">`resid` Должен сопоставляться с `id` атрибутом `Url` элемента в `Resources` элементе.</span><span class="sxs-lookup"><span data-stu-id="b20ae-122">The `resid` must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |

## <a name="see-also"></a><span data-ttu-id="b20ae-123">См. также</span><span class="sxs-lookup"><span data-stu-id="b20ae-123">See also</span></span>

- [<span data-ttu-id="b20ae-124">Runtimes</span><span class="sxs-lookup"><span data-stu-id="b20ae-124">Runtimes</span></span>](runtimes.md)
