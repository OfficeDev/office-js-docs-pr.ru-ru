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
# <a name="tokens-element"></a>Элемент tokens

Определяет маркеры, которые можно использовать в URL-адресах шаблонов.

**Тип надстройки:** надстройки области задач

## <a name="syntax"></a>Синтаксис

```XML
<Tokens></Tokens>
```

## <a name="contained-in"></a>Содержится в

[екстендедоверридес](extendedoverrides.md)

## <a name="must-contain"></a>Должен содержать

|Элемент|Контентная|Почта|Область задач|
|:-----|:-----|:-----|:-----|
|[Маркер](token.md)|||x|

## <a name="example"></a>Пример

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