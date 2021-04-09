---
title: Время запуска в файле манифеста
description: Элемент Runtime настраивает надстройку для использования общего времени запуска JavaScript для различных компонентов, например ленты, области задач, пользовательских функций.
ms.date: 04/08/2021
localization_priority: Normal
ms.openlocfilehash: fa95608d7eff57d68b96ef5b04ec9d33ee63f173
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652246"
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

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **resid**  |  Да  | Указывает расположение URL-адреса страницы HTML для надстройки. Символ может быть не более 32 символов и должен соответствовать `resid` `id` атрибуту `Url` элемента `Resources` элемента. |
|  **срок службы**  |  Нет  | Значение по умолчанию является и не нужно `lifetime` `short` задано. Надстройки Outlook используют только `short` значение. Если вы хотите использовать совместное время работы в надстройки Excel, заочная настройка значения `long` . |

## <a name="see-also"></a>См. также

- [Runtimes](runtimes.md)
- [Настройка надстройки Office для использования общей среды выполнения JavaScript](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Настройка надстройки Outlook для активации на основе событий](../../outlook/autolaunch.md)
