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
# <a name="token-element"></a>Элемент Token

Определяет отдельный маркер URL-адреса.

**Тип надстройки:** надстройки области задач

## <a name="syntax"></a>Синтаксис

```XML
<Token Name="string" DefaultValue="string" xsi:type=["LocaleToken" | "RequirementsToken"] ></Token>
```

## <a name="contained-in"></a>Содержится в

[Обнаружения](tokens.md)

## <a name="can-contain"></a>Может содержать

|Элемент|Контентная|Почта|Область задач|
|:-----|:-----|:-----|:-----|
|[Override](override.md)|||x|

## <a name="attributes"></a>Атрибуты

|Атрибут|Описание|
|:-----|:-----|
|DefaultValue|Значение по умолчанию для этого маркера, если ни одно условие не соответствует ни одному из дочерних `<Override>` элементов.|
|Имя|Имя маркера. Это имя определяется пользователем. Тип маркера определяется атрибутом Type.|
|xsi:type|Определяет тип маркера. Для этого атрибута необходимо задать один из следующих параметров:  `"RequirementsToken"` или  `"LocaleToken"` .|

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