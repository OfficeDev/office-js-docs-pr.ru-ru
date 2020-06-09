---
title: Элемент SourceLocation в файле манифеста
description: Элемент SourceLocation указывает расположение исходных файлов для надстройки Office.
ms.date: 05/12/2020
localization_priority: Normal
ms.openlocfilehash: 9af2337263314bec5ce04eb0d22626ab368c19ef
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608728"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="738a8-103">Элемент SourceLocation</span><span class="sxs-lookup"><span data-stu-id="738a8-103">SourceLocation element</span></span>

<span data-ttu-id="738a8-104">Указывает расположение исходных файлов для надстройки Office в виде URL-адреса длиной от 1 до 2018 символов.</span><span class="sxs-lookup"><span data-stu-id="738a8-104">Specifies the source file locations for your Office Add-in as a URL between 1 and 2018 characters long.</span></span> <span data-ttu-id="738a8-105">В качестве источника необходимо указать адрес HTTPS, а не путь к файлу.</span><span class="sxs-lookup"><span data-stu-id="738a8-105">The source location must be an HTTPS address, not a file path.</span></span>

<span data-ttu-id="738a8-106">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="738a8-106">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="738a8-107">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="738a8-107">Syntax</span></span>

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a><span data-ttu-id="738a8-108">Содержится в</span><span class="sxs-lookup"><span data-stu-id="738a8-108">Contained in</span></span>

- <span data-ttu-id="738a8-109">[DefaultSettings](defaultsettings.md) (надстройки области задач и контентные надстройки)</span><span class="sxs-lookup"><span data-stu-id="738a8-109">[DefaultSettings](defaultsettings.md) (Content and task pane add-ins)</span></span>
- <span data-ttu-id="738a8-110">[FormSettings](formsettings.md) (почтовые надстройки)</span><span class="sxs-lookup"><span data-stu-id="738a8-110">[FormSettings](formsettings.md) (Mail add-ins)</span></span>
- <span data-ttu-id="738a8-111">[ExtensionPoint](extensionpoint.md) (контекстные и лаунчевент (Предварительная версия) почтовые надстройки)</span><span class="sxs-lookup"><span data-stu-id="738a8-111">[ExtensionPoint](extensionpoint.md) (Contextual and LaunchEvent (preview) mail add-ins)</span></span>

## <a name="can-contain"></a><span data-ttu-id="738a8-112">Может содержать</span><span class="sxs-lookup"><span data-stu-id="738a8-112">Can contain</span></span>

[<span data-ttu-id="738a8-113">Override</span><span class="sxs-lookup"><span data-stu-id="738a8-113">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="738a8-114">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="738a8-114">Attributes</span></span>

|<span data-ttu-id="738a8-115">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="738a8-115">**Attribute**</span></span>|<span data-ttu-id="738a8-116">**Тип**</span><span class="sxs-lookup"><span data-stu-id="738a8-116">**Type**</span></span>|<span data-ttu-id="738a8-117">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="738a8-117">**Required**</span></span>|<span data-ttu-id="738a8-118">**Описание**</span><span class="sxs-lookup"><span data-stu-id="738a8-118">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="738a8-119">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="738a8-119">DefaultValue</span></span>|<span data-ttu-id="738a8-120">URL-адрес</span><span class="sxs-lookup"><span data-stu-id="738a8-120">URL</span></span>|<span data-ttu-id="738a8-121">Обязательный</span><span class="sxs-lookup"><span data-stu-id="738a8-121">required</span></span>|<span data-ttu-id="738a8-122">Задает значение этого параметра по умолчанию для языкового стандарта, указанного в элементе [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="738a8-122">Specifies the default value for this setting for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
