---
title: Runtimes in the manifest file
description: Элемент Runtimes указывает времени работы надстройки.
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: afbcc6a909c51d2ed56292ef1541193f7f698d28
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789165"
---
# <a name="runtimes-element"></a>Элемент Runtimes

Указывает времени работы надстройки. Child of the [`<Host>`](host.md) element.

> [!NOTE]
> При запуске в Office для Windows надстройка использует браузер Internet Explorer 11.

В Excel этот элемент позволяет ленте, области задач и пользовательским функциям использовать ту же времени работы. Дополнительные сведения см. в настройках надстройки Excel для использования общей времени [работы JavaScript.](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)

В Outlook этот элемент включает активацию надстройки на основе событий. Дополнительные сведения см. в настройке [надстройки Outlook для активации на основе событий.](../../outlook/autolaunch.md)

**Тип надстройки:** Области задач, почта

> [!IMPORTANT]
> **Outlook**: функция активации на [](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) основе событий в настоящее время находится в предварительной версии и доступна только в Outlook в Интернете. Дополнительные сведения см. в [предварительном просмотре функции активации на основе событий.](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)

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
| [Время выполнения](runtime.md) | Да |  Время работы надстройки. |

## <a name="see-also"></a>См. также

- [Время выполнения](runtime.md)
