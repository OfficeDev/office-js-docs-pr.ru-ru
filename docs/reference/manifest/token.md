---
title: Элемент token в файле манифеста
description: Указывает маркер или подстановочный знак, который можно использовать с шаблонами URL-адресов в манифесте.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 5e26af44c566ab09ac81c8194e1ae7d85aaac327
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996704"
---
# <a name="token-element"></a><span data-ttu-id="cf21c-103">Элемент Token</span><span class="sxs-lookup"><span data-stu-id="cf21c-103">Token element</span></span>

<span data-ttu-id="cf21c-104">Определяет отдельный маркер URL-адреса.</span><span class="sxs-lookup"><span data-stu-id="cf21c-104">Defines an individual URL token.</span></span>

<span data-ttu-id="cf21c-105">**Тип надстройки:** надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="cf21c-105">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="cf21c-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="cf21c-106">Syntax</span></span>

```XML
<Token Name="string" DefaultValue="string" xsi:type=["LocaleToken" | "RequirementsToken"] ></Token>
```

## <a name="contained-in"></a><span data-ttu-id="cf21c-107">Содержится в</span><span class="sxs-lookup"><span data-stu-id="cf21c-107">Contained in</span></span>

[<span data-ttu-id="cf21c-108">Обнаружения</span><span class="sxs-lookup"><span data-stu-id="cf21c-108">Tokens</span></span>](tokens.md)

## <a name="can-contain"></a><span data-ttu-id="cf21c-109">Может содержать</span><span class="sxs-lookup"><span data-stu-id="cf21c-109">Can contain</span></span>

|<span data-ttu-id="cf21c-110">Элемент</span><span class="sxs-lookup"><span data-stu-id="cf21c-110">Element</span></span>|<span data-ttu-id="cf21c-111">Контентная</span><span class="sxs-lookup"><span data-stu-id="cf21c-111">Content</span></span>|<span data-ttu-id="cf21c-112">Почта</span><span class="sxs-lookup"><span data-stu-id="cf21c-112">Mail</span></span>|<span data-ttu-id="cf21c-113">Область задач</span><span class="sxs-lookup"><span data-stu-id="cf21c-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="cf21c-114">Override</span><span class="sxs-lookup"><span data-stu-id="cf21c-114">Override</span></span>](override.md)|||<span data-ttu-id="cf21c-115">x</span><span class="sxs-lookup"><span data-stu-id="cf21c-115">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="cf21c-116">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="cf21c-116">Attributes</span></span>

|<span data-ttu-id="cf21c-117">Атрибут</span><span class="sxs-lookup"><span data-stu-id="cf21c-117">Attribute</span></span>|<span data-ttu-id="cf21c-118">Описание</span><span class="sxs-lookup"><span data-stu-id="cf21c-118">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="cf21c-119">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="cf21c-119">DefaultValue</span></span>|<span data-ttu-id="cf21c-120">Значение по умолчанию для этого маркера, если ни одно условие не соответствует ни одному из дочерних `<Override>` элементов.</span><span class="sxs-lookup"><span data-stu-id="cf21c-120">Default value for this token if no condition in any child `<Override>` element matches.</span></span>|
|<span data-ttu-id="cf21c-121">Имя</span><span class="sxs-lookup"><span data-stu-id="cf21c-121">Name</span></span>|<span data-ttu-id="cf21c-122">Имя маркера.</span><span class="sxs-lookup"><span data-stu-id="cf21c-122">Token name.</span></span> <span data-ttu-id="cf21c-123">Это имя определяется пользователем.</span><span class="sxs-lookup"><span data-stu-id="cf21c-123">This name is user-defined.</span></span> <span data-ttu-id="cf21c-124">Тип маркера определяется атрибутом Type.</span><span class="sxs-lookup"><span data-stu-id="cf21c-124">The type of the token is determined by the type attribute.</span></span>|
|<span data-ttu-id="cf21c-125">xsi:type</span><span class="sxs-lookup"><span data-stu-id="cf21c-125">xsi:type</span></span>|<span data-ttu-id="cf21c-126">Определяет тип маркера.</span><span class="sxs-lookup"><span data-stu-id="cf21c-126">Defines the kind of Token.</span></span> <span data-ttu-id="cf21c-127">Для этого атрибута необходимо задать один из следующих параметров:  `"RequirementsToken"` или  `"LocaleToken"` .</span><span class="sxs-lookup"><span data-stu-id="cf21c-127">This attribute should be set to one of:  `"RequirementsToken"`,  or  `"LocaleToken"`.</span></span>|

## <a name="example"></a><span data-ttu-id="cf21c-128">Пример</span><span class="sxs-lookup"><span data-stu-id="cf21c-128">Example</span></span>

```XML
<OfficeApp ...>
  <!-- other elements omitted -->
  <ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.locale}/extended-manifest-overrides.json">
    <Tokens>
      <Token Name="locale" DefaultValue="en-us" xsi:type="LocaleToken">
        <Override Locale="es-*" Value="es-es" />
        <Override Locale="es-mx" Value="es-mx" />
        <Override Locale="fr-*" Value="fr-fr" />
        <Override Locale="ja-jp" Value="ja-jp" />
      </Token>
    <Tokens>
  </ExtendedOverrides>
</OfficeApp>
```