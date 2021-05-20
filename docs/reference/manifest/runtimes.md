---
title: Время времени времени времени времени в файле манифеста
description: Элемент Runtimes определяет время выполнения надстройки.
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: 80336674c6d954bb9e0c6892feb41cb2f03c5859
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555299"
---
# <a name="runtimes-element"></a>Элемент времени бегут

Определяет время выполнения надстройки. Дитя [`<Host>`](host.md) элемента.

> [!NOTE]
> При запуске Office на Windows, надстройку, которая имеет элемент в манифесте, не обязательно работает в том же `<Runtimes>` контроле веб-вида, как это было бы в противном случае. Для получения дополнительной информации о том, как Windows и Office определяют, какой элемент управления веб-видом обычно [используется, Office см.](../../concepts/browsers-used-by-office-web-add-ins.md) Если описанные там условия для использования Microsoft Edge с WebView2 (Chromium-based) выполнены, то надстройок использует этот браузер независимо от того, есть ли у него `<Runtimes>` элемент. Однако, когда эти условия не выполнены, надстройа с `<Runtimes>` элементом всегда использует Internet Explorer 11 независимо от Windows или Microsoft 365 версии.

**Тип дополнения:** Панель задач, Почта

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
| [Время выполнения](runtime.md) | Да |  Время выполнения надстройок. **Важно**: В настоящее время можно определить только один `<Runtime>` элемент. |

## <a name="see-also"></a>См. также

- [Время выполнения](runtime.md)
- [Настройка надстройки Office для использования общей среды выполнения JavaScript](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Настройте Outlook для активации на основе событий](../../outlook/autolaunch.md)
