---
title: Отламывка надстроек с помощью средств разработчика для устаревшая версия Microsoft Edge
description: Отламывка надстроек с помощью средств разработчика в устаревшая версия Microsoft Edge.
ms.date: 11/02/2021
ms.localizationpriority: medium
ms.openlocfilehash: 62f27e2ee266e3b6a92d090e8008b74bac4a9663
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744682"
---
# <a name="debug-add-ins-using-developer-tools-in-microsoft-edge-legacy"></a>Отламывка надстроек с помощью средств разработчика в устаревшая версия Microsoft Edge

В этой статье показано, как отлагировать клиентский код (JavaScript или TypeScript) надстройки при условии, при которых будут выполнены следующие условия.

- Вы не можете или не хотите отлаговка с помощью инструментов, встроенных в ваш IDE; или вы столкнулись с проблемой, которая возникает только при запуске надстройки за пределами IDE.
- На вашем компьютере используется сочетание Windows и Office, использующих оригинальный контроль веб-просмотров Edge Edge, EdgeHTML.

> [!TIP]
> Сведения о отладке с помощью edge Legacy в Visual Studio Code см. в Microsoft Office расширения надстройки для [Visual Studio Code](debug-with-vs-extension.md).

Чтобы определить, какой браузер вы используете, см. в Office [надстройки](../concepts/browsers-used-by-office-web-add-ins.md). 

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

> [!NOTE]
> Чтобы установить версию Office с устаревшим веб-Office Edge или заставить текущую версию Office использовать Edge Legacy, см. в странице [Switch to the Edge Legacy webview](#switch-to-the-edge-legacy-webview).

## <a name="debug-a-task-pane-add-in-using-microsoft-edge-devtools-preview"></a>Отладка надстройки области задач с помощью Microsoft Edge DevTools Preview

1. Установите [предварительный Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?activetab=pivot%3Aoverviewtab). (Слово "Preview" находится в названии по историческим причинам. Более последней версии нет.)

   > [!NOTE]
   > Если надстройка имеет команду [](../design/add-in-commands.md) надстройки, которая выполняет функцию, функция выполняется в скрытом процессе браузера, который Microsoft Edge DevTools не могут обнаружить или прикрепить, поэтому описанный в этой статье метод не может использоваться для отладки кода в функции.

1. [Боковая](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) загрузка и запуск надстройки.
1. Запустите Microsoft Edge DevTools.
1. Перейдите на вкладку **Локальные**. Имя вашей надстройки будет указано в списке. (На вкладке отображаются только процессы, запущенные в EdgeHTML. Средство не может присоединяться к процессам, запущенным в других браузерах или веб-Microsoft Edge (WebView2) и Internet Explorer (Trident).)

   :::image type="content" source="../images/edge-devtools-with-add-in-process.png" alt-text="Снимок экрана Edge DevTools, показывающий процесс с именем отладки с устаревшими краями.":::

1. Выберите имя надстройки, чтобы открыть его в средствах.
1. Перейдите на вкладку **Отладчик**.
1. Откройте файл, который необходимо отламыть следующими шагами.

   1. На панели задач отладки выберите **Показать поиск в файлах**. Это откроет окно поиска.
   1. Введите строку кода из файла, который необходимо отлагировать в поле поиска. Это должно быть что-то, что вряд ли будет в любом другом файле.
   1. Выберите кнопку обновления.
   1. В результатах поиска выберите строку, чтобы открыть файл кода на области выше результатов поиска.

   :::image type="content" source="../images/open-file-in-edge-devtools.png" alt-text="Снимок экрана вкладки отладки Edge DevTools с 4 частями с меткой A через D.":::

1. Чтобы установить точку разрыва, выберите строку в файле кода. Точка разрыва регистрируется в области **стек** вызовов (внизу справа). Кроме того, в файле кода может быть красная точка, но это не выглядит надежно.
1. Выполните функции в надстройке, необходимые для срабатывания точки останова.

> [!TIP]
> Дополнительные сведения об использовании средств см. в [Microsoft Edge (EdgeHTML).](/archive/microsoft-edge/legacy/developer/devtools-guide/)

## <a name="debug-a-dialog-in-an-add-in"></a>Отламыв диалоговое окно надстройки

Если надстройка использует API Office диалогов, диалоговое окно запускается в отдельном процессе от области задач (если таковое имеется), и средства должны присоединяться к этому процессу. Выполните указанные ниже действия.

1. Запустите надстройку и средства.
1. Откройте диалоговое окно и выберите кнопку **Обновить** в средствах. Показан диалоговое окно. Его имя происходит от элемента `<title>` HTML-файла, открытого в диалоговом окте.
1. Выберите процесс, чтобы открыть его и отладить так же, как описано в разделе Отладка надстройки области задач с помощью [Microsoft Edge DevTools Preview](#debug-a-task-pane-add-in-using-microsoft-edge-devtools-preview).

   :::image type="content" source="../images/edge-devtools-with-add-in-and-dialog-processes.png" alt-text="Снимок экрана Edge DevTools, показывающий процесс с именем My Dialog.":::

## <a name="switch-to-the-edge-legacy-webview"></a>Переключиться на веб-просмотр Edge Legacy

Существует два способа переключения веб-приложения Edge Legacy. Вы можете запустить простую команду в командной подсказке или установить версию Office, использующую Edge Legacy по умолчанию. Рекомендуем первый метод. Но второй вариант следует использовать в следующих сценариях.

- Ваш проект был разработан с Visual Studio и IIS. Это не node.js основе.
- Вы хотите быть абсолютно надежным в тестировании.
- Если по какой-либо причине средство командной строки не работает.

### <a name="switch-via-the-command-line"></a>Переключение через командную строку

[!INCLUDE [Steps to switch browsers with the command line tool](../includes/use-legacy-edge-or-ie.md)]

### <a name="install-a-version-of-office-that-uses-edge-legacy"></a>Установка версии Office с использованием Edge Legacy

[!INCLUDE [Steps to install Office that uses Edge Legacy or Internet Explorer](../includes/install-office-that-uses-legacy-edge-or-ie.md)]
