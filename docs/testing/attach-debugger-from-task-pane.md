# <a name="attach-a-debugger-from-the-task-pane"></a>Подключение отладчика из области задач

В Office 2016 для Windows (сборка 77xx.xxxx или более поздней версии) можно подключать отладчик из области задач. Функция "Подключить отладчик" подключит отладчик непосредственно к нужному процессу Internet Explorer. Вы можете подключить отладчик независимо от того, какой инструмент используете: генератор Yeoman, Visual Studio Code, node.js, Angular или другой. 

Для запуска средства **подключения отладчика** откройте меню **Личные данные** в правом верхнем углу области задач (выделено красным на рисунке ниже).   

 >  **Примечания**.  
   - В настоящее время единственный поддерживаемый инструмент отладки — [Visual Studio 2015](https://www.visualstudio.com/downloads/) с [обновлением 3](https://msdn.microsoft.com/en-us/library/mt752379.aspx) или более поздней версии. Если у вас нет Visual Studio, при выборе команды **Подключить отладчик** не будет выполняться никаких действий.   
   - Для отладки клиентского кода JavaScript можно использовать только средство **Подключить отладчик**. Для отладки серверного кода, например на сервере Node.js, существует множество вариантов. Сведения о том, как выполнять отладку в Visual Studio Code, см. в статье [Отладка Node.js в VS Code](https://code.visualstudio.com/docs/nodejs/nodejs-debugging). Если вы не используете Visual Studio Code, выполните поиск по запросу "отладка Node.js" или "отладка {имя_сервера}".

![Снимок экрана: меню подключения отладчика](../images/attach-debugger.png)

Выберите **Подключить отладчик**. Откроется диалоговое окно **JIT-отладчик Visual Studio** (см. рисунок ниже). 

![Снимок экрана: JIT-отладчик Visual Studio](../images/visual-studio-debugger.png)

В **обозревателе решений** Visual Studio вы увидите файлы кода.   Вы можете задать точки останова для отлаживаемой строки кода в Visual Studio.

Дополнительные сведения об отладке в Visual Studio см. в следующих статьях:

-   Дополнительные сведения о запуске и использовании Проводника DOM в Visual Studio приведены в совете № 4 в разделе [Советы и рекомендации](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates/#tips_tricks) записи в блоге [Создание отличных приложений для Office с помощью новых шаблонов проекта](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates).
-   Как задать точки останова, можно узнать в статье [Использование точек останова](https://msdn.microsoft.com/en-US/library/5557y8b4.aspx).
-   Сведения об использовании F12 см. в статье [Использование средств разработчика F12](https://msdn.microsoft.com/en-us/library/bg182326(v=vs.85).aspx).

## <a name="additional-resources"></a>Дополнительные ресурсы

- [Создание и отладка надстроек Office в Visual Studio](../../docs/get-started/create-and-debug-office-add-ins-in-visual-studio.md)
- [Создание надстройки Office с помощью любого редактора](../../docs/get-started/create-an-office-add-in-using-any-editor.md)
- [Публикация надстройки Office](../publish/publish.md)
