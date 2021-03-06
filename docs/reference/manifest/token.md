---
title: Элемент маркера в файле манифеста
description: Указывает маркер или под диктовую карточку, которые можно использовать с шаблонами URL-адресов в манифесте.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 48078f8211a8fd3f0e3f9d7c3f3aabd1d31b0a6d
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505369"
---
# <a name="token-element"></a><span data-ttu-id="78e16-103">Элемент Token</span><span class="sxs-lookup"><span data-stu-id="78e16-103">Token element</span></span>

<span data-ttu-id="78e16-104">Определяет отдельный маркер URL-адреса.</span><span class="sxs-lookup"><span data-stu-id="78e16-104">Defines an individual URL token.</span></span> <span data-ttu-id="78e16-105">Дополнительные сведения об использовании этого элемента см. в дополнительных сведениях о работе с расширенными [переопределениями манифеста.](../../develop/extended-overrides.md)</span><span class="sxs-lookup"><span data-stu-id="78e16-105">For more information about the use of this element, see [Work with extended overrides of the manifest](../../develop/extended-overrides.md).</span></span>

<span data-ttu-id="78e16-106">**Тип надстройки:** надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="78e16-106">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="78e16-107">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="78e16-107">Syntax</span></span>

```XML
<Token Name="string" DefaultValue="string" xsi:type=["LocaleToken" | "RequirementsToken"] ></Token>
```

## <a name="contained-in"></a><span data-ttu-id="78e16-108">Содержится в</span><span class="sxs-lookup"><span data-stu-id="78e16-108">Contained in</span></span>

[<span data-ttu-id="78e16-109">Tokens</span><span class="sxs-lookup"><span data-stu-id="78e16-109">Tokens</span></span>](tokens.md)

## <a name="can-contain"></a><span data-ttu-id="78e16-110">Может содержать</span><span class="sxs-lookup"><span data-stu-id="78e16-110">Can contain</span></span>

|<span data-ttu-id="78e16-111">Элемент</span><span class="sxs-lookup"><span data-stu-id="78e16-111">Element</span></span>|<span data-ttu-id="78e16-112">Контентная</span><span class="sxs-lookup"><span data-stu-id="78e16-112">Content</span></span>|<span data-ttu-id="78e16-113">Почта</span><span class="sxs-lookup"><span data-stu-id="78e16-113">Mail</span></span>|<span data-ttu-id="78e16-114">Область задач</span><span class="sxs-lookup"><span data-stu-id="78e16-114">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="78e16-115">Override</span><span class="sxs-lookup"><span data-stu-id="78e16-115">Override</span></span>](override.md)|||<span data-ttu-id="78e16-116">x</span><span class="sxs-lookup"><span data-stu-id="78e16-116">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="78e16-117">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="78e16-117">Attributes</span></span>

|<span data-ttu-id="78e16-118">Атрибут</span><span class="sxs-lookup"><span data-stu-id="78e16-118">Attribute</span></span>|<span data-ttu-id="78e16-119">Описание</span><span class="sxs-lookup"><span data-stu-id="78e16-119">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="78e16-120">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="78e16-120">DefaultValue</span></span>|<span data-ttu-id="78e16-121">Значение по умолчанию для этого маркера, если условие в любом `<Override>` детском элементе не совпадает.</span><span class="sxs-lookup"><span data-stu-id="78e16-121">Default value for this token if no condition in any child `<Override>` element matches.</span></span>|
|<span data-ttu-id="78e16-122">Имя</span><span class="sxs-lookup"><span data-stu-id="78e16-122">Name</span></span>|<span data-ttu-id="78e16-123">Имя маркера.</span><span class="sxs-lookup"><span data-stu-id="78e16-123">Token name.</span></span> <span data-ttu-id="78e16-124">Это имя определяется пользователем.</span><span class="sxs-lookup"><span data-stu-id="78e16-124">This name is user-defined.</span></span> <span data-ttu-id="78e16-125">Тип маркера определяется атрибутом типа.</span><span class="sxs-lookup"><span data-stu-id="78e16-125">The type of the token is determined by the type attribute.</span></span>|
|<span data-ttu-id="78e16-126">xsi:type</span><span class="sxs-lookup"><span data-stu-id="78e16-126">xsi:type</span></span>|<span data-ttu-id="78e16-127">Определяет тип Маркера.</span><span class="sxs-lookup"><span data-stu-id="78e16-127">Defines the kind of Token.</span></span> <span data-ttu-id="78e16-128">Этот атрибут должен быть заданной для одного из:  `"RequirementsToken"` или  `"LocaleToken"` .</span><span class="sxs-lookup"><span data-stu-id="78e16-128">This attribute should be set to one of:  `"RequirementsToken"`,  or  `"LocaleToken"`.</span></span>|

## <a name="example"></a><span data-ttu-id="78e16-129">Пример</span><span class="sxs-lookup"><span data-stu-id="78e16-129">Example</span></span>

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