---
title: Время запуска в файле манифеста
description: Элемент Runtimes указывает время работы надстройки.
ms.date: 04/08/2021
localization_priority: Normal
ms.openlocfilehash: a5cd05a0890615375bf3466caf70d22f9912d951
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652238"
---
# <a name="runtimes-element"></a>Элемент Runtimes

Указывает время запуска надстройки. Ребенок [`<Host>`](host.md) элемента.

> [!NOTE]
> При работе в Office на Windows надстройка использует браузер Internet Explorer 11.

**Тип надстройки:** Области задач, Почта

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a>Синтаксис

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>Содержится в

[Host](host.md)

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
| [Runtime](runtime.md) | Да |  Время запуска надстройки. |

## <a name="see-also"></a>См. также

- [Runtime](runtime.md)
- [Настройка надстройки Office для использования общей среды выполнения JavaScript](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Настройка надстройки Outlook для активации на основе событий](../../outlook/autolaunch.md)
