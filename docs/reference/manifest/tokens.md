---
title: Элемент Маркеры в файле манифеста
description: Указывает маркеры или под диктовки, которые можно использовать с URL-шаблонами в манифесте.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 5d42abab46ecc6e7ab465144f061d26da52c0eb3e2623acd8a8a2912ecc13312
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57095790"
---
# <a name="tokens-element"></a>Элемент Маркеры

Определяет маркеры, которые можно использовать в URL-адресах шаблонов. Дополнительные сведения об использовании этого элемента см. в дополнительных сведениях о работе с расширенными [переопределениями манифеста.](../../develop/extended-overrides.md)

**Тип надстройки:** надстройки области задач

## <a name="syntax"></a>Синтаксис

```XML
<Tokens></Tokens>
```

## <a name="contained-in"></a>Содержится в

[ExtendedOverrides](extendedoverrides.md)

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