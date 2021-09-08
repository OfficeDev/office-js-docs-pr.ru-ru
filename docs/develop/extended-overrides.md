---
title: Работа с расширенными переопределениями манифеста
description: Узнайте, как настроить функции расширяемости с расширенными переопределениями манифеста.
ms.date: 02/23/2021
localization_priority: Normal
ms.openlocfilehash: 4eb8936e8a01b81a3883f848446d20ebf4ecf863
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938635"
---
# <a name="work-with-extended-overrides-of-the-manifest"></a>Работа с расширенными переопределениями манифеста

Некоторые функции extensibility Office настраиваются с JSON-файлами, которые находятся на сервере, а не с XML-манифестом надстройки.

> [!NOTE]
> В этой статье предполагается, что вы знакомы с Office манифестами надстройки и их ролью в надстройки. Пожалуйста, [Office XML-манифест](add-in-manifests.md)надстройки, если вы еще не были недавно.

В следующей таблице указаны функции расширяемости, которые требуют расширенного переопределения, а также ссылки на документацию по этой функции.

| Функция | Инструкции по разработке |
| :----- | :----- |
| Сочетания клавиш | [Добавление ярлыков настраиваемой клавиатуры в Office надстройки](../design/keyboard-shortcuts.md) |

Схема, определяемая форматом JSON, имеет [схему расширенного манифеста.](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)

> [!TIP]
> Эта статья несколько абстрактна. Рассмотрите чтение одной из статей в таблице, чтобы добавить ясность к понятиям.

## <a name="tell-office-where-to-find-the-json-file"></a>Сообщите Office, где найти файл JSON

Используйте манифест, чтобы Office, где найти файл JSON. Сразу *ниже* (не внутри) элемента `<VersionOverrides>` манифеста добавьте элемент [ExtendedOverrides.](../reference/manifest/extendedoverrides.md) Установите атрибут `Url` для полного URL-адреса файла JSON. Ниже приводится пример простейшего `<ExtendedOverrides>` элемента.

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json"></ExtendedOverrides>
</OfficeApp>
```

Ниже приводится пример очень простого расширенного переопределения JSON-файла. Он назначает клавишу ярлык CTRL+SHIFT+A функции (определенной в другом месте), которая открывает области задач надстройки.

```json
{
    "actions": [
        {
            "id": "SHOWTASKPANE",
            "type": "ExecuteFunction",
            "name": "Show task pane for add-in"
        }
    ],
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "CTRL+SHIFT+A"
            }
        }
    ]
}
```

## <a name="localize-the-extended-overrides-file"></a>Локализация расширенного переопределения файла

Если надстройка поддерживает несколько локальных элементов, можно использовать атрибут элемента, чтобы указать Office файлу `ResourceUrl` `<ExtendedOverrides>` локализованных ресурсов. Ниже приведен пример.

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json" 
                       ResourceUrl="https://contoso.com/addin/my-resources.json">
    </ExtendedOverrides>
</OfficeApp>
```

Дополнительные сведения о том, как создавать и использовать файл ресурсов, как ссылаться на его ресурсы в расширенном файле переопределения, а также дополнительные параметры, не рассмотренные здесь, см. в материале [Localize extended overrides.](localization.md#localize-extended-overrides)
