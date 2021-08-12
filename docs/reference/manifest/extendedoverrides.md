---
title: Элемент ExtendedOverrides в файле манифеста
description: Указывает URL-адреса для расширения манифеста в формате JSON.
ms.date: 02/23/2021
localization_priority: Normal
ms.openlocfilehash: f2b9ea409763119b5bec5286ecdc5f15c94c49e6312a13209197e6457353f369
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57083588"
---
# <a name="extendedoverrides-element"></a>Элемент ExtendedOverrides

Указывает полные URL-адреса для файлов в формате JSON, которые расширяют манифест. Подробные сведения об использовании этого элемента и его потомкных элементов см. в см. в описании [Work with extended overrides of the manifest.](../../develop/extended-overrides.md)

**Тип надстройки:** надстройки области задач

## <a name="syntax"></a>Синтаксис

```XML
<ExtendedOverrides Url="string" [ResourcesUrl="string"] ></ExtendedOverrides>
```

## <a name="contained-in"></a>Содержится в

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Может содержать

|Элемент|Контентная|Почта|Область задач|
|:-----|:-----|:-----|:-----|
|[Tokens](tokens.md)|||x|

## <a name="attributes"></a>Атрибуты

|Атрибут|Описание|
|:-----|:-----|
|Url (обязательно)| Полный URL-адрес расширенного файла JSON переопределяется. В будущем это значение может быть url-шаблоном, использующим маркеры, определенные элементом [Tokens.](tokens.md) См. [примеры](#examples).|
|ResourcesUrl (необязательный) | Полный URL-адрес файла, который предоставляет дополнительные ресурсы, например локализованные строки, для файла, указанного в `Url` атрибуте. Это может быть URL-шаблон, использующий маркеры, определенные элементом [Tokens.](tokens.md)|

## <a name="examples"></a>Примеры

```XML
<OfficeApp ...>
  <!-- other elements omitted -->
  <ExtendedOverrides Url="http://contoso.com/addinmetadata/extended-manifest-overrides.json"
                     ResourceUrl="https://contoso.com/addin/my-resources.json">
  </ExtendedOverrides>
</OfficeApp>
```

В будущем это значение может быть url-шаблоном, использующим маркеры, определенные элементом [Tokens.](tokens.md) Ниже приведен пример.

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
