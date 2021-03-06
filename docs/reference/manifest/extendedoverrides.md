---
title: Элемент ExtendedOverrides в файле манифеста
description: Указывает URL-адреса для расширения манифеста в формате JSON.
ms.date: 02/23/2021
localization_priority: Normal
ms.openlocfilehash: f433c9c5604f3fae35580ba20780ea6fe91401c7
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505474"
---
# <a name="extendedoverrides-element"></a><span data-ttu-id="33379-103">Элемент ExtendedOverrides</span><span class="sxs-lookup"><span data-stu-id="33379-103">ExtendedOverrides element</span></span>

<span data-ttu-id="33379-104">Указывает полные URL-адреса для файлов в формате JSON, которые расширяют манифест.</span><span class="sxs-lookup"><span data-stu-id="33379-104">Specifies the full URLs for JSON-formatted files that extend the manifest.</span></span> <span data-ttu-id="33379-105">Подробные сведения об использовании этого элемента и его потомкных элементов см. в см. в описании [Work with extended overrides of the manifest.](../../develop/extended-overrides.md)</span><span class="sxs-lookup"><span data-stu-id="33379-105">For detailed information about the use of this element and its descendent elements, see [Work with extended overrides of the manifest](../../develop/extended-overrides.md).</span></span>

<span data-ttu-id="33379-106">**Тип надстройки:** надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="33379-106">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="33379-107">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="33379-107">Syntax</span></span>

```XML
<ExtendedOverrides Url="string" [ResourcesUrl="string"] ></ExtendedOverrides>
```

## <a name="contained-in"></a><span data-ttu-id="33379-108">Содержится в</span><span class="sxs-lookup"><span data-stu-id="33379-108">Contained in</span></span>

[<span data-ttu-id="33379-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="33379-109">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="33379-110">Может содержать</span><span class="sxs-lookup"><span data-stu-id="33379-110">Can contain</span></span>

|<span data-ttu-id="33379-111">Элемент</span><span class="sxs-lookup"><span data-stu-id="33379-111">Element</span></span>|<span data-ttu-id="33379-112">Контентная</span><span class="sxs-lookup"><span data-stu-id="33379-112">Content</span></span>|<span data-ttu-id="33379-113">Почта</span><span class="sxs-lookup"><span data-stu-id="33379-113">Mail</span></span>|<span data-ttu-id="33379-114">Область задач</span><span class="sxs-lookup"><span data-stu-id="33379-114">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="33379-115">Tokens</span><span class="sxs-lookup"><span data-stu-id="33379-115">Tokens</span></span>](tokens.md)|||<span data-ttu-id="33379-116">x</span><span class="sxs-lookup"><span data-stu-id="33379-116">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="33379-117">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="33379-117">Attributes</span></span>

|<span data-ttu-id="33379-118">Атрибут</span><span class="sxs-lookup"><span data-stu-id="33379-118">Attribute</span></span>|<span data-ttu-id="33379-119">Описание</span><span class="sxs-lookup"><span data-stu-id="33379-119">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="33379-120">Url (обязательно)</span><span class="sxs-lookup"><span data-stu-id="33379-120">Url (required)</span></span>| <span data-ttu-id="33379-121">Полный URL-адрес расширенного файла JSON переопределяется.</span><span class="sxs-lookup"><span data-stu-id="33379-121">The full URL of the extended overrides JSON file.</span></span> <span data-ttu-id="33379-122">В будущем это значение может быть url-шаблоном, использующим маркеры, определенные элементом [Tokens.](tokens.md)</span><span class="sxs-lookup"><span data-stu-id="33379-122">In the future, this value could be a URL template that uses tokens defined by the [Tokens](tokens.md) element.</span></span> <span data-ttu-id="33379-123">См. [примеры](#examples).</span><span class="sxs-lookup"><span data-stu-id="33379-123">See [Examples](#examples).</span></span>|
|<span data-ttu-id="33379-124">ResourcesUrl (необязательный)</span><span class="sxs-lookup"><span data-stu-id="33379-124">ResourcesUrl (optional)</span></span> | <span data-ttu-id="33379-125">Полный URL-адрес файла, который предоставляет дополнительные ресурсы, например локализованные строки, для файла, указанного в `Url` атрибуте.</span><span class="sxs-lookup"><span data-stu-id="33379-125">The full URL of a file that provides supplemental resources, such as localized strings, for the file specified in the `Url` attribute.</span></span> <span data-ttu-id="33379-126">Это может быть URL-шаблон, использующий маркеры, определенные элементом [Tokens.](tokens.md)</span><span class="sxs-lookup"><span data-stu-id="33379-126">This could be a URL template that uses tokens defined by the [Tokens](tokens.md) element.</span></span>|

## <a name="examples"></a><span data-ttu-id="33379-127">Примеры</span><span class="sxs-lookup"><span data-stu-id="33379-127">Examples</span></span>

```XML
<OfficeApp ...>
  <!-- other elements omitted -->
  <ExtendedOverrides Url="http://contoso.com/addinmetadata/extended-manifest-overrides.json"
                     ResourceUrl="https://contoso.com/addin/my-resources.json">
  </ExtendedOverrides>
</OfficeApp>
```

<span data-ttu-id="33379-128">В будущем это значение может быть url-шаблоном, использующим маркеры, определенные элементом [Tokens.](tokens.md)</span><span class="sxs-lookup"><span data-stu-id="33379-128">In the future, this value could be a URL template that uses tokens defined by the [Tokens](tokens.md) element.</span></span> <span data-ttu-id="33379-129">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="33379-129">The following is an example.</span></span>

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
