---
title: Время запуска в файле манифеста
description: Элемент Runtimes указывает время работы надстройки.
ms.date: 09/28/2021
ms.localizationpriority: medium
ms.openlocfilehash: 758bb7b830009d6691190a0279440a52da724624
ms.sourcegitcommit: 489befc41e543a4fb3c504fd9b3f61322134c1ef
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/06/2021
ms.locfileid: "60138606"
---
# <a name="runtimes-element"></a>Элемент Runtimes

Указывает время запуска надстройки. Ребенок [`<Host>`](host.md) элемента.

> [!NOTE]
> При работе Office на Windows, надстройка с элементом в манифесте не обязательно будет работать в том же элементе управления веб-просмотром, что и `<Runtimes>` в противном случае. Дополнительные сведения о том, как версии Windows и Office, которые обычно используются для управления [веб-просмотром,](../../concepts/browsers-used-by-office-web-add-ins.md)см. в Office надстройки. Если условия, описанные в нем для Microsoft Edge с webView2 (Chromium на основе), будут выполнены, то надстройка использует этот браузер независимо от того, имеет ли он `<Runtimes>` элемент. Однако, если эти условия не выполнены, надстройка с элементом всегда использует Internet Explorer 11 независимо от Windows или `<Runtimes>` Microsoft 365 версии.

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

[Host](host.md)

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
| [Runtime](runtime.md) | Да |  Время запуска надстройки. **Важно.** В настоящее время можно определить только один `<Runtime>` элемент. |

## <a name="see-also"></a>См. также

- [Runtime](runtime.md)
- [Настройка надстройки Office для использования общей среды выполнения JavaScript](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Настройка надстройки Outlook для активации на основе событий](../../outlook/autolaunch.md)
