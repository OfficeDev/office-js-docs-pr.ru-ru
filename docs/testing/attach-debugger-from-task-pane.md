---
title: Подключение отладчика из области задач
description: ''
ms.date: 08/06/2019
localization_priority: Priority
ms.openlocfilehash: 4433eccc165d616c7b8ebff428e1426c4b98be49
ms.sourcegitcommit: dc78ee2a89fe3d4cd6f748be1eec9081c1077502
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2019
ms.locfileid: "36231307"
---
# <a name="attach-a-debugger-from-the-task-pane"></a>Подключение отладчика из области задач

В Office 2016 для Windows (сборка 77xx.xxxx или более поздней версии) можно подключать отладчик из области задач. Функция "Подключить отладчик" подключит отладчик непосредственно к нужному процессу Internet Explorer. Вы можете подключить отладчик независимо от того, какой инструмент используете: генератор Yeoman, Visual Studio Code, Node.js, Angular или другой. 

Для запуска средства **подключения отладчика** откройте меню **Личные данные** в правом верхнем углу области задач (выделено красным на рисунке ниже).   

> [!NOTE]
> - В настоящее время поддерживается только отладчик [Visual Studio 2015](https://www.visualstudio.com/downloads/) с [обновлением 3](https://msdn.microsoft.com/library/mt752379.aspx) или более поздней версии. Если у вас нет Visual Studio, выбор параметра **Подключить отладчик** не даст результата.   
> - Для отладки клиентского кода JavaScript можно использовать только средство **Подключить отладчик**. Для отладки серверного кода, например на сервере Node.js, существует множество вариантов. Сведения о том, как выполнять отладку в Visual Studio Code, см. в статье [Отладка Node.js в VS Code](https://code.visualstudio.com/docs/nodejs/nodejs-debugging). Если вы не используете Visual Studio Code, выполните поиск по запросу "отладка Node.js" или "отладка {имя_сервера}".

![Снимок экрана: меню подключения отладчика](../images/attach-debugger.png)

Выберите элемент **Подключить отладчик**. Откроется диалоговое окно **JIT-отладчик Visual Studio** (см. рисунок ниже). 

![Снимок экрана: JIT-отладчик Visual Studio](../images/visual-studio-debugger.png)

В **обозревателе решений** Visual Studio вы увидите файлы кода.   Вы можете задать точки останова для отлаживаемой строки кода в Visual Studio.

> [!NOTE]
> Если меню "Личные данные" не отображается, отладить надстройку можно с помощью Visual Studio. Убедитесь, что надстройка области задач открыта в Office, и выполните указанные ниже действия.
>
> 1. В Visual Studio выберите **ОТЛАДКА** > **Присоединиться к процессу**.
> 2. В окне **Присоединение к процессу** выберите все доступные процессы Iexplore.exe, а затем нажмите кнопку **Присоединиться**.

Дополнительные сведения об отладке в Visual Studio см. в следующих статьях:

-   Дополнительные сведения о запуске и использовании Проводника DOM в Visual Studio приведены в совете № 4 в разделе [Советы и рекомендации](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates/#tips_tricks) записи в блоге [Создание отличных приложений для Office с помощью новых шаблонов проекта](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates).
-   Как задать точки останова, можно узнать в статье [Использование точек останова](/visualstudio/debugger/using-breakpoints?view=vs-2015).
-   Сведения об использовании F12 см. в статье [Использование средств разработчика F12](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85)).

## <a name="see-also"></a>См. также

- [Создание и отладка надстроек Office в Visual Studio](../develop/create-and-debug-office-add-ins-in-visual-studio.md)
- [Публикация надстройки Office](../publish/publish.md)
