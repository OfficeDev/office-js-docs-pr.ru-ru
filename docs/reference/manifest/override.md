---
title: Элемент Override в файле манифеста
description: Элемент override позволяет указать значение параметра для дополнительного языкового стандарта.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 139a4089a36d8a8adfa71d4a0947b02f5b163b52
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641454"
---
# <a name="override-element"></a>Элемент Override

Предоставляет способ указать значение параметра для дополнительного языкового стандарта.

**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач

## <a name="syntax"></a>Синтаксис

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a>Содержится в

|Элемент|
|:-----|
|[CitationText](citationtext.md)|
|[Описание](description.md)|
|[DictionaryName](dictionaryname.md)|
|[DictionaryHomePage](dictionaryhomepage.md)|
|[DisplayName](displayname.md)|
|[HighResolutionIconUrl](highresolutioniconurl.md)|
|[IconUrl](iconurl.md)|
|[QueryUri](queryuri.md)|
|[SourceLocation](sourcelocation.md)|
|[SupportUrl](supporturl.md)|

## <a name="attributes"></a>Атрибуты

|Атрибут|Тип|Обязательный|Описание|
|:-----|:-----|:-----|:-----|
|Языковой стандарт|string|Обязательный|Задает имя языка и региональных параметров для языкового стандарта этого переопределения в формате языковых тегов BCP 47, например `"en-US"`.|
|Значение|string|Обязательный|Задает значение параметра, представленное для указанного языкового стандарта.|

## <a name="see-also"></a>См. также

- [Локализация надстроек для Office](../../develop/localization.md)
