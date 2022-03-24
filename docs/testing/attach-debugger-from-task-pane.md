---
title: Подключение отладчика из области задач
description: Узнайте, как прикрепить отладку из области задач.
ms.date: 01/27/2022
ms.localizationpriority: medium
ms.openlocfilehash: 0363b7966ab3da11167cb4b0cd324df28fd9efb3
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744754"
---
# <a name="attach-a-debugger-from-the-task-pane"></a>Подключение отладчика из области задач

В некоторых средах отладка может быть присоединена к уже запущенной надстройки Office надстройки. Это может быть полезно при отлаговлении надстройки, которая уже находится в постановке или производстве. Если надстройка еще разрабатывается и тестируется, см. в обзоре отладки Office [надстроек](debug-add-ins-overview.md).

Описанный в этой статье метод можно использовать только при следующих условиях.

- Надстройка работает в Office на Windows.
- Компьютер использует сочетание версий Windows и Office, использующих управление веб-Chromium edge (Chromium веб-просмотров) WebView2. Чтобы определить, какой браузер вы используете, см. в Office [надстройки](../concepts/browsers-used-by-office-web-add-ins.md).

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

Чтобы запустить отладку, выберите верхний правый угол области задач, чтобы активировать меню **Personality** (как показано на красном круге на следующем изображении).

![Снимок экрана меню Attach Debugger.](../images/attach-debugger.png)

Выберите **Attach Debugger**. В этом случае запускается Microsoft Edge (Chromium на основе) средств разработчика. Используйте средства, описанные в надстройках Debug, с помощью средств разработчика в [Microsoft Edge (Chromium основе)](debug-add-ins-using-devtools-edge-chromium.md).

## <a name="see-also"></a>См. также

- [Обзор отладки надстроек Office](debug-add-ins-overview.md)
