---
title: Среды выполнения в файле манифеста
description: Элемент Runtimes указывает среду выполнения надстройки.
ms.date: 05/29/2020
localization_priority: Normal
ms.openlocfilehash: a8598a8f926e6d6905c147f5c554f1d40a692ad9
ms.sourcegitcommit: 09a8683ff29cf06d0d1d822be83cf0798f1ccdf9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/01/2020
ms.locfileid: "44471326"
---
# <a name="runtimes-element"></a>Элемент среды выполнения

Задает среду выполнения надстройки. Дочерний [`<Host>`](host.md) элемент. Если `Runtimes` элемент присутствует в манифесте, надстройка по умолчанию будет использовать браузер Internet Explorer 11.

В Excel этот элемент позволяет использовать одну и ту же среду выполнения для ленты, области задач и пользовательских функций. Дополнительные сведения можно найти в статье [Настройка надстройки Excel для использования общей среды выполнения JavaScript](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).

В Outlook этот элемент включает активацию надстройки на основе событий. Дополнительные сведения см. [в разделе Настройка надстройки Outlook для активации на основе событий](../../outlook/autolaunch.md).

**Тип надстройки:** Область задач, почта

> [!IMPORTANT]
> **Excel**: общая среда выполнения в настоящее время доступна только в Excel для Windows.
>
> **Outlook**: функция активации на основе событий в настоящее время находится [в предварительной версии](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) и доступна только в Outlook в Интернете. Дополнительные сведения см. [в статье Просмотр функции активации на основе событий](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).

## <a name="syntax"></a>Синтаксис

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>Содержится в

[Host](host.md) (Узел)

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
| [Время выполнения](runtime.md) | Да |  Среда выполнения надстройки. |

## <a name="see-also"></a>См. также

- [Время выполнения](runtime.md)
