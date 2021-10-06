---
title: Время запуска в файле манифеста
description: Элемент Runtime настраивает надстройку для использования общего времени запуска JavaScript для различных компонентов, например ленты, области задач, пользовательских функций.
ms.date: 09/28/2021
ms.localizationpriority: medium
ms.openlocfilehash: acdff8f7ffb1e9392c1671eadc36a79348ece5fa
ms.sourcegitcommit: 489befc41e543a4fb3c504fd9b3f61322134c1ef
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/06/2021
ms.locfileid: "60138445"
---
# <a name="runtime-element"></a>Элемент runtime

Настраивает надстройку для использования общего времени запуска JavaScript, чтобы все компоненты запускались в одном и том же времени. Ребенок [`<Runtimes>`](runtimes.md) элемента.

**Тип надстройки:** Области задач, Почта

**Допустимо только в этих схемах VersionOverrides:**

 - Области задач 1.0
 - Почта 1.1

Дополнительные сведения см. в [манифесте "Версия переопределения".](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)

**Связанные с этими наборами требований:**

- [SharedRuntime 1.1](../requirement-sets/shared-runtime-requirement-sets.md) (Только при ее использования в надстройке области задач.)

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
