---
title: Время выполнения в файле манифеста
description: Элемент Runtime настраивает надстройки для использования общего времени выполнения JavaScript для различных компонентов, например ленты, панели задач, пользовательских функций.
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: c59e5a23e53940aea46c758d710b4a455cb5c0cc
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555306"
---
# <a name="runtime-element"></a>Элемент времени выполнения

Настраивает надстройки для использования общего времени выполнения JavaScript, чтобы все различные компоненты запускаются в одно и то же время выполнения. Дитя [`<Runtimes>`](runtimes.md) элемента.

**Тип дополнения:** Панель задач, Почта

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
| [Переопределение](override.md) (предварительный просмотр) | Нет | **Outlook**: Определяет местоположение URL-адреса файла JavaScript, который требуется Outlook Desktop [для обработчиков токов точки расширения LaunchEvent.](../../reference/manifest/extensionpoint.md#launchevent-preview) **Важно:** В настоящее время вы можете определить только `<Override>` один элемент, и он должен быть типа `javascript` .|

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **resid**  |  Да  | Определяет местоположение URL страницы HTML для надстройки. Может `resid` быть не более 32 символов и должен `id` соответствовать атрибуту `Url` элемента в `Resources` элементе. |
|  **продолжительность жизни**  |  Нет  | Значение по `lifetime` умолчанию `short` для является и не должно быть указано. Outlook надстройки используют только `short` значение. Если вы хотите использовать общее время выполнения в Excel в дополнение, явно установите `long` значение. |

## <a name="see-also"></a>См. также

- [Runtimes](runtimes.md)
- [Настройка надстройки Office для использования общей среды выполнения JavaScript](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Настройте Outlook для активации на основе событий](../../outlook/autolaunch.md)
