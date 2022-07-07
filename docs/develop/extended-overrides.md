---
title: Работа с расширенными переопределениями манифеста
description: Узнайте, как настроить функции расширяемости с помощью расширенных переопределений манифеста.
ms.date: 02/23/2021
ms.localizationpriority: medium
ms.openlocfilehash: 43e9820f54f2812130f7f86529c52b20b92811a0
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/06/2022
ms.locfileid: "66659955"
---
# <a name="work-with-extended-overrides-of-the-manifest"></a>Работа с расширенными переопределениями манифеста

Некоторые функции расширяемости надстроек Office настраиваются с помощью JSON-файлов, размещенных на сервере, а не XML-манифеста надстройки.

> [!NOTE]
> В этой статье предполагается, что вы знакомы с манифестами надстроек Office и их ролью в надстройке. Прочитайте [XML-манифест](add-in-manifests.md) надстроек Office, если вы еще этого не делать.

В следующей таблице указаны функции расширяемости, для которых требуется расширенное переопределение, а также ссылки на документацию по этой функции.

| Возможность | Инструкции по разработке |
| :----- | :----- |
| Сочетания клавиш | [Добавление настраиваемых сочетаний клавиш в надстройки Office](../design/keyboard-shortcuts.md) |

Схема, которая определяет формат JSON, [является схемой extended-manifest](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

> [!TIP]
> Эта статья является в некоторой степени абстрактной. Рассмотрите возможность чтения одной из статей в таблице, чтобы внести ясности в основные понятия.

## <a name="tell-office-where-to-find-the-json-file"></a>Укажите Office, где найти JSON-файл

Используйте манифест, чтобы сообщить Office, где найти JSON-файл. Непосредственно *под* элементом манифеста (не внутри) **\<VersionOverrides\>** добавьте [элемент ExtendedOverrides](/javascript/api/manifest/extendedoverrides) . Задайте `Url` для атрибута полный URL-адрес JSON-файла. Ниже приведен пример простейшего из возможных элементов **\<ExtendedOverrides\>** .

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json"></ExtendedOverrides>
</OfficeApp>
```

Ниже приведен пример очень простого расширенного переопределения JSON-файла. Он назначает сочетание клавиш CTRL+SHIFT+A функции (определенной в другом месте), которая открывает область задач надстройки.

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

## <a name="localize-the-extended-overrides-file"></a>Локализация файла расширенных переопределений

Если надстройка поддерживает несколько языковых стандартов, `ResourceUrl` **\<ExtendedOverrides\>** можно использовать атрибут элемента, чтобы указать Office на файл локализованных ресурсов. Ниже приведен пример.

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json" 
                       ResourceUrl="https://contoso.com/addin/my-resources.json">
    </ExtendedOverrides>
</OfficeApp>
```

Дополнительные сведения о том, как создать и использовать файл ресурсов, как ссылаться на его ресурсы в файле расширенных переопределений и дополнительные параметры, не описанные здесь, см. в разделе ["](localization.md#localize-extended-overrides)Локализация расширенных переопределений".
