---
title: Время запуска в файле манифеста
description: Элемент Runtimes указывает время работы надстройки.
ms.date: 04/16/2021
localization_priority: Normal
ms.openlocfilehash: 8f4a602c05b9af7bde9f644ef40b61a214e66cd5
ms.sourcegitcommit: da8ad214406f2e1cd80982af8a13090e76187dbd
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/21/2021
ms.locfileid: "51917088"
---
# <a name="runtimes-element"></a>Элемент Runtimes

Указывает время запуска надстройки. Ребенок [`<Host>`](host.md) элемента.

> [!NOTE]
> При работе в Office on Windows надстройка с элементом манифеста не обязательно будет работать в том же элементе управления веб-просмотром, что и `<Runtimes>` в противном случае. Дополнительные сведения о том, как версии Windows и Office определяют, как обычно используется управление веб-просмотром, см. в браузерах, используемых [надстройки Office.](../../concepts/browsers-used-by-office-web-add-ins.md) Если условия, описанные там для использования Microsoft Edge с WebView2 (на основе хрома), выполнены, то надстройка использует этот браузер независимо от того, имеет ли он `<Runtimes>` элемент. Однако, если эти условия не выполнены, надстройка с элементом всегда использует Internet Explorer 11 независимо от версии Windows или `<Runtimes>` Microsoft 365.

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
| [Время выполнения](runtime.md) | Да |  Время запуска надстройки. |

## <a name="see-also"></a>См. также

- [Время выполнения](runtime.md)
- [Настройка надстройки Office для использования общей среды выполнения JavaScript](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Настройка надстройки Outlook для активации на основе событий](../../outlook/autolaunch.md)
