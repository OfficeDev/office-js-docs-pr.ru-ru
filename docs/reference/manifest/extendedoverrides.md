---
title: Элемент Екстендедоверридес в файле манифеста
description: Задает URL-адреса для расширения манифеста в формате JSON.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 76491af34d1caf0ec266826df97a5363e336b85d
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996708"
---
# <a name="extendedoverrides-element"></a><span data-ttu-id="f89e5-103">Элемент Екстендедоверридес</span><span class="sxs-lookup"><span data-stu-id="f89e5-103">ExtendedOverrides element</span></span>

<span data-ttu-id="f89e5-104">Задает полные URL-адреса для файлов в формате JSON, которые расширяют манифест.</span><span class="sxs-lookup"><span data-stu-id="f89e5-104">Specifies the full URLs for JSON-formatted files that extend the manifest.</span></span>

<span data-ttu-id="f89e5-105">**Тип надстройки:** надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="f89e5-105">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="f89e5-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="f89e5-106">Syntax</span></span>

```XML
<ExtendedOverrides Url="string" [ResourcesUrl="string"] ></ExtendedOverrides>
```

## <a name="contained-in"></a><span data-ttu-id="f89e5-107">Содержится в</span><span class="sxs-lookup"><span data-stu-id="f89e5-107">Contained in</span></span>

[<span data-ttu-id="f89e5-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="f89e5-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="f89e5-109">Может содержать</span><span class="sxs-lookup"><span data-stu-id="f89e5-109">Can contain</span></span>

|<span data-ttu-id="f89e5-110">Элемент</span><span class="sxs-lookup"><span data-stu-id="f89e5-110">Element</span></span>|<span data-ttu-id="f89e5-111">Контентная</span><span class="sxs-lookup"><span data-stu-id="f89e5-111">Content</span></span>|<span data-ttu-id="f89e5-112">Почта</span><span class="sxs-lookup"><span data-stu-id="f89e5-112">Mail</span></span>|<span data-ttu-id="f89e5-113">Область задач</span><span class="sxs-lookup"><span data-stu-id="f89e5-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="f89e5-114">Обнаружения</span><span class="sxs-lookup"><span data-stu-id="f89e5-114">Tokens</span></span>](tokens.md)|||<span data-ttu-id="f89e5-115">x</span><span class="sxs-lookup"><span data-stu-id="f89e5-115">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="f89e5-116">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="f89e5-116">Attributes</span></span>

|<span data-ttu-id="f89e5-117">Атрибут</span><span class="sxs-lookup"><span data-stu-id="f89e5-117">Attribute</span></span>|<span data-ttu-id="f89e5-118">Описание</span><span class="sxs-lookup"><span data-stu-id="f89e5-118">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="f89e5-119">URL-адрес (обязательный)</span><span class="sxs-lookup"><span data-stu-id="f89e5-119">Url (required)</span></span>| <span data-ttu-id="f89e5-120">Полный URL-адрес расширенных переопределений JSON-файла.</span><span class="sxs-lookup"><span data-stu-id="f89e5-120">The full URL of the extended overrides JSON file.</span></span> <span data-ttu-id="f89e5-121">Это может быть шаблон URL-адреса, в котором используются маркеры, определенные элементом [tokens](tokens.md) .</span><span class="sxs-lookup"><span data-stu-id="f89e5-121">This could be a URL template that uses tokens defined by the [Tokens](tokens.md) element.</span></span>|
|<span data-ttu-id="f89e5-122">Ресаурцесурл (необязательно)</span><span class="sxs-lookup"><span data-stu-id="f89e5-122">ResourcesUrl (optional)</span></span> | <span data-ttu-id="f89e5-123">Полный URL-адрес файла, который предоставляет дополнительные ресурсы, такие как локализованные строки, для файла, указанного в `Url` атрибуте.</span><span class="sxs-lookup"><span data-stu-id="f89e5-123">The full URL of a file that provides supplemental resources, such as localized strings, for the file specified in the `Url` attribute.</span></span> <span data-ttu-id="f89e5-124">Это может быть шаблон URL-адреса, в котором используются маркеры, определенные элементом [tokens](tokens.md) .</span><span class="sxs-lookup"><span data-stu-id="f89e5-124">This could be a URL template that uses tokens defined by the [Tokens](tokens.md) element.</span></span>|

## <a name="example"></a><span data-ttu-id="f89e5-125">Пример</span><span class="sxs-lookup"><span data-stu-id="f89e5-125">Example</span></span>

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
