---
title: Элемент tokens в файле манифеста
description: Задает маркеры или подстановочные знаки, которые можно использовать с шаблонами URL-адресов в манифесте.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: a50de7c2c3e8ebeb9425c1677a94bbcc62281d3b
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996698"
---
# <a name="tokens-element"></a><span data-ttu-id="a9311-103">Элемент tokens</span><span class="sxs-lookup"><span data-stu-id="a9311-103">Tokens element</span></span>

<span data-ttu-id="a9311-104">Определяет маркеры, которые можно использовать в URL-адресах шаблонов.</span><span class="sxs-lookup"><span data-stu-id="a9311-104">Defines tokens that could be used in template URLs.</span></span>

<span data-ttu-id="a9311-105">**Тип надстройки:** надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="a9311-105">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="a9311-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="a9311-106">Syntax</span></span>

```XML
<Tokens></Tokens>
```

## <a name="contained-in"></a><span data-ttu-id="a9311-107">Содержится в</span><span class="sxs-lookup"><span data-stu-id="a9311-107">Contained in</span></span>

[<span data-ttu-id="a9311-108">екстендедоверридес</span><span class="sxs-lookup"><span data-stu-id="a9311-108">ExtendedOverrides</span></span>](extendedoverrides.md)

## <a name="must-contain"></a><span data-ttu-id="a9311-109">Должен содержать</span><span class="sxs-lookup"><span data-stu-id="a9311-109">Must contain</span></span>

|<span data-ttu-id="a9311-110">Элемент</span><span class="sxs-lookup"><span data-stu-id="a9311-110">Element</span></span>|<span data-ttu-id="a9311-111">Контентная</span><span class="sxs-lookup"><span data-stu-id="a9311-111">Content</span></span>|<span data-ttu-id="a9311-112">Почта</span><span class="sxs-lookup"><span data-stu-id="a9311-112">Mail</span></span>|<span data-ttu-id="a9311-113">Область задач</span><span class="sxs-lookup"><span data-stu-id="a9311-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="a9311-114">Маркер</span><span class="sxs-lookup"><span data-stu-id="a9311-114">Token</span></span>](token.md)|||<span data-ttu-id="a9311-115">x</span><span class="sxs-lookup"><span data-stu-id="a9311-115">x</span></span>|

## <a name="example"></a><span data-ttu-id="a9311-116">Пример</span><span class="sxs-lookup"><span data-stu-id="a9311-116">Example</span></span>

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