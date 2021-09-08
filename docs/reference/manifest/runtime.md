---
title: Время запуска в файле манифеста
description: Элемент Runtime настраивает надстройку для использования общего времени запуска JavaScript для различных компонентов, например ленты, области задач, пользовательских функций.
ms.date: 05/19/2021
localization_priority: Normal
ms.openlocfilehash: cd09abe31ff57eac629c6c61c873c5c886f73f9c
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937334"
---
# <a name="runtime-element"></a>Элемент runtime

Настраивает надстройку для использования общего времени запуска JavaScript, чтобы все компоненты запускались в одном и том же времени. Ребенок [`<Runtimes>`](runtimes.md) элемента.

**Тип надстройки:** Области задач, Почта

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a>Синтаксис

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>Содержится в

- [Runtimes](runtimes.md)

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
| [Override](override.md) | Нет | **Outlook:** указывает расположение URL-адреса файла JavaScript, который Outlook для обработчиков точеки [расширения LaunchEvent.](../../reference/manifest/extensionpoint.md#launchevent) **Важно:** в настоящее время можно определить только один элемент и `<Override>` он должен быть типа `javascript` .|

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **resid**  |  Да  | Указывает расположение URL-адреса страницы HTML для надстройки. Символ может быть не более 32 символов и должен соответствовать `resid` `id` атрибуту `Url` элемента `Resources` элемента. |
|  **срок службы**  |  Нет  | Значение по умолчанию является и не нужно `lifetime` `short` задано. Outlook надстройки используют только `short` значение. Если вы хотите использовать совместное время работы в Excel надстройки, явно установите значение `long` . |

## <a name="see-also"></a>См. также

- [Runtimes](runtimes.md)
- [Настройка надстройки Office для использования общей среды выполнения JavaScript](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Настройка надстройки Outlook для активации на основе событий](../../outlook/autolaunch.md)
