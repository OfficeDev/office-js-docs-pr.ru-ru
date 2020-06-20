---
title: Подключение отладчика из области задач
description: Сведения о том, как подключить отладчик из области задач
ms.date: 06/17/2020
localization_priority: Normal
ms.openlocfilehash: 53cfce211241dbdf3d16e8a126e059a2f2db3f23
ms.sourcegitcommit: b939312ffdeb6e0a0dfe085db7efe0ff143ef873
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/19/2020
ms.locfileid: "44810844"
---
# <a name="attach-a-debugger-from-the-task-pane"></a>Подключение отладчика из области задач

In Office 2016 on Windows, Build 77xx.xxxx or later, you can attach the debugger from the task pane. The attach debugger feature will directly attach the debugger to the correct Internet Explorer process for you. You can attach a debugger regardless of whether you are using Yeoman Generator, Visual Studio Code, Node.js, Angular, or another tool. 

Для запуска средства **подключения отладчика** откройте меню **Личные данные** в правом верхнем углу области задач (выделено красным на рисунке ниже).   

> [!NOTE]
> - В настоящее время поддерживается только отладчик [Visual Studio 2015](https://www.visualstudio.com/downloads/) с [обновлением 3](https://msdn.microsoft.com/library/mt752379.aspx) или более поздней версии. Если вы не установили Visual Studio, то при выборе параметра " **Присоединение отладчика** " не будет выполняться никаких действий.   
> - You can only debug client-side JavaScript with the **Attach Debugger** tool. To debug server-side code, such as with a Node.js server, you have many options. For information on how to debug with Visual Studio Code, see [Node.js Debugging in VS Code](https://code.visualstudio.com/docs/nodejs/nodejs-debugging). If you are not using Visual Studio Code, search for "debug Node.js" or "debug {name-of-server}".

![Снимок экрана: меню подключения отладчика](../images/attach-debugger.png)

Select **Attach Debugger**. This launches the **Visual Studio Just-in-Time Debugger** dialog box, as shown in the following image. 

![Снимок экрана: JIT-отладчик Visual Studio](../images/visual-studio-debugger.png)

In Visual Studio, you will see the code files in **Solution Explorer**.   You can set breakpoints to the line of code you want to debug in Visual Studio.

> [!NOTE]
> Если меню "Личные данные" не отображается, отладить надстройку можно с помощью Visual Studio. Убедитесь, что надстройка области задач открыта в Office, и выполните указанные ниже действия.
>
> 1. В Visual Studio выберите **ОТЛАДКА** > **Присоединиться к процессу**.
> 2. В разделе **Доступные процессы** выберите *либо* все доступные процессы `Iexplore.exe`, *либо* все доступные процессы `MicrosoftEdge*.exe`, в зависимости от того, [использует ли ваша надстройка Internet Explorer или Microsoft Edge](../concepts/browsers-used-by-office-web-add-ins.md), а затем нажмите кнопку **Присоединиться**.

Дополнительные сведения об отладке в Visual Studio см. в следующих статьях:

-    Дополнительные сведения о запуске и использовании Проводника DOM в Visual Studio приведены в совете № 4 в разделе [Советы и рекомендации](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates/#tips_tricks) записи в блоге [Создание отличных приложений для Office с помощью новых шаблонов проекта](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates).
-    Как задать точки останова, можно узнать в статье [Использование точек останова](/visualstudio/debugger/using-breakpoints?view=vs-2015).
-    Сведения об использовании F12 см. в статье [Использование средств разработчика F12](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85)).
-   Сведения об использовании средств разработчика в Microsoft Edge см. на странице [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?activetab=pivot%3Aoverviewtab).

## <a name="see-also"></a>См. также

- [Отладка надстроек Office в Visual Studio](../develop/debug-office-add-ins-in-visual-studio.md)
- [Публикация надстройки Office](../publish/publish.md)
- [Расширение отладчика надстроек Microsoft Office для Visual Studio Code](debug-with-vs-extension.md)