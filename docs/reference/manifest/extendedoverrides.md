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
# <a name="extendedoverrides-element"></a>Элемент Екстендедоверридес

Задает полные URL-адреса для файлов в формате JSON, которые расширяют манифест.

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
|[Обнаружения](tokens.md)|||x|

## <a name="attributes"></a>Атрибуты

|Атрибут|Описание|
|:-----|:-----|
|URL-адрес (обязательный)| Полный URL-адрес расширенных переопределений JSON-файла. Это может быть шаблон URL-адреса, в котором используются маркеры, определенные элементом [tokens](tokens.md) .|
|Ресаурцесурл (необязательно) | Полный URL-адрес файла, который предоставляет дополнительные ресурсы, такие как локализованные строки, для файла, указанного в `Url` атрибуте. Это может быть шаблон URL-адреса, в котором используются маркеры, определенные элементом [tokens](tokens.md) .|

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
