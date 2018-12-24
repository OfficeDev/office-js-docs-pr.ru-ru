---
title: Элемент SourceLocation в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: dc432ebb9482e8e9b8be5d90a838357ccf519ad3
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433518"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="9c3f9-102">Элемент SourceLocation</span><span class="sxs-lookup"><span data-stu-id="9c3f9-102">SourceLocation element</span></span>

<span data-ttu-id="9c3f9-p101">Указывает расположения исходного файла для надстройки Office как URL-адреса длиной от 1 до 2018 символов. В качестве источника необходимо указать адрес HTTPS, а не путь к файлу.</span><span class="sxs-lookup"><span data-stu-id="9c3f9-p101">Specifies the source file location(s) for your Office Add-in as a URL between 1 and 2018 characters long. The source location must be an HTTPS address, not a file path.</span></span>

<span data-ttu-id="9c3f9-105">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="9c3f9-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="9c3f9-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="9c3f9-106">Syntax</span></span>

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a><span data-ttu-id="9c3f9-107">Содержится в</span><span class="sxs-lookup"><span data-stu-id="9c3f9-107">Contained in</span></span>

- <span data-ttu-id="9c3f9-108">[DefaultSettings](defaultsettings.md) (надстройки области задач и контентные надстройки)</span><span class="sxs-lookup"><span data-stu-id="9c3f9-108">[DefaultSettings](defaultsettings.md) (Content and task pane add-ins)</span></span>
- <span data-ttu-id="9c3f9-109">[FormSettings](formsettings.md) (почтовые надстройки)</span><span class="sxs-lookup"><span data-stu-id="9c3f9-109">[FormSettings](formsettings.md) (Mail add-ins)</span></span>
- <span data-ttu-id="9c3f9-110">[ExtensionPoint](extensionpoint.md) (контекстные почтовые надстройки)</span><span class="sxs-lookup"><span data-stu-id="9c3f9-110">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins)</span></span>

## <a name="can-contain"></a><span data-ttu-id="9c3f9-111">Может содержать</span><span class="sxs-lookup"><span data-stu-id="9c3f9-111">Can contain</span></span>

[<span data-ttu-id="9c3f9-112">Переопределение</span><span class="sxs-lookup"><span data-stu-id="9c3f9-112">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="9c3f9-113">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="9c3f9-113">Attributes</span></span>

|<span data-ttu-id="9c3f9-114">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="9c3f9-114">**Attribute**</span></span>|<span data-ttu-id="9c3f9-115">**Тип**</span><span class="sxs-lookup"><span data-stu-id="9c3f9-115">**Type**</span></span>|<span data-ttu-id="9c3f9-116">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="9c3f9-116">**Required**</span></span>|<span data-ttu-id="9c3f9-117">**Описание**</span><span class="sxs-lookup"><span data-stu-id="9c3f9-117">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="9c3f9-118">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="9c3f9-118">DefaultValue</span></span>|<span data-ttu-id="9c3f9-119">URL-адрес</span><span class="sxs-lookup"><span data-stu-id="9c3f9-119">URL</span></span>|<span data-ttu-id="9c3f9-120">Обязательный</span><span class="sxs-lookup"><span data-stu-id="9c3f9-120">required</span></span>|<span data-ttu-id="9c3f9-121">Задает значение этого параметра по умолчанию для языкового стандарта, указанного в элементе [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="9c3f9-121">Specifies the default value for this setting for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
