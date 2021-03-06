---
title: Элемент Маркеры в файле манифеста
description: Указывает маркеры или под диктовки, которые можно использовать с URL-шаблонами в манифесте.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 8680b985068c44e93f601a2b24e2f28899eb483d
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505327"
---
# <a name="tokens-element"></a><span data-ttu-id="6a32f-103">Элемент Маркеры</span><span class="sxs-lookup"><span data-stu-id="6a32f-103">Tokens element</span></span>

<span data-ttu-id="6a32f-104">Определяет маркеры, которые можно использовать в URL-адресах шаблонов.</span><span class="sxs-lookup"><span data-stu-id="6a32f-104">Defines tokens that could be used in template URLs.</span></span> <span data-ttu-id="6a32f-105">Дополнительные сведения об использовании этого элемента см. в дополнительных сведениях о работе с расширенными [переопределениями манифеста.](../../develop/extended-overrides.md)</span><span class="sxs-lookup"><span data-stu-id="6a32f-105">For more information about the use of this element, see [Work with extended overrides of the manifest](../../develop/extended-overrides.md).</span></span>

<span data-ttu-id="6a32f-106">**Тип надстройки:** надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="6a32f-106">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="6a32f-107">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="6a32f-107">Syntax</span></span>

```XML
<Tokens></Tokens>
```

## <a name="contained-in"></a><span data-ttu-id="6a32f-108">Содержится в</span><span class="sxs-lookup"><span data-stu-id="6a32f-108">Contained in</span></span>

[<span data-ttu-id="6a32f-109">ExtendedOverrides</span><span class="sxs-lookup"><span data-stu-id="6a32f-109">ExtendedOverrides</span></span>](extendedoverrides.md)

## <a name="must-contain"></a><span data-ttu-id="6a32f-110">Должен содержать</span><span class="sxs-lookup"><span data-stu-id="6a32f-110">Must contain</span></span>

|<span data-ttu-id="6a32f-111">Элемент</span><span class="sxs-lookup"><span data-stu-id="6a32f-111">Element</span></span>|<span data-ttu-id="6a32f-112">Контентная</span><span class="sxs-lookup"><span data-stu-id="6a32f-112">Content</span></span>|<span data-ttu-id="6a32f-113">Почта</span><span class="sxs-lookup"><span data-stu-id="6a32f-113">Mail</span></span>|<span data-ttu-id="6a32f-114">Область задач</span><span class="sxs-lookup"><span data-stu-id="6a32f-114">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="6a32f-115">Маркер</span><span class="sxs-lookup"><span data-stu-id="6a32f-115">Token</span></span>](token.md)|||<span data-ttu-id="6a32f-116">x</span><span class="sxs-lookup"><span data-stu-id="6a32f-116">x</span></span>|

## <a name="example"></a><span data-ttu-id="6a32f-117">Пример</span><span class="sxs-lookup"><span data-stu-id="6a32f-117">Example</span></span>

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