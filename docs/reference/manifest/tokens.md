---
title: Элемент Маркеры в файле манифеста
description: Указывает маркеры или под диктовки, которые можно использовать с URL-шаблонами в манифесте.
ms.date: 11/06/2020
ms.localizationpriority: medium
ms.openlocfilehash: 3e52543bdb53709ea005f63a3a990650905d70cd
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154691"
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