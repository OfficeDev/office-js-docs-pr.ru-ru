Если пространство имен `CONTOSO` недоступно в меню автозаполнения, для регистрации надстройки в Excel выполните следующие действия.

### <a name="excel-on-windows-or-mac"></a>[Excel для Windows или Mac](#tab/excel-windows)

1. В Excel выберите вкладку **Вставка**, а затем нажмите стрелку вниз, находящуюся справа от элемента **Мои надстройки**.

    :::image type="content" source="../images/select-insert.png" alt-text="Снимок экрана: лента &quot;Вставка&quot; в Excel для Windows с выделенной стрелкой &quot;Мои надстройки&quot;":::

1. В списке доступных надстроек найдите раздел **Надстройки разработчика** и выберите вашу надстройку **starcount**, чтобы ее зарегистрировать.

    :::image type="content" source="../images/list-starcount.png" alt-text="Снимок экрана: лента &quot;Вставка&quot; в Excel для Windows с выделенной надстройкой &quot;Пользовательские функции Excel&quot; в списке &quot;Мои надстройки&quot;.":::

# <a name="excel-on-the-web"></a>[Excel в Интернете](#tab/excel-online)

1. В Excel на вкладке **Вставка** выберите пункт **Надстройки**.

    :::image type="content" source="../images/excel-cf-online-register-add-in-1.png" alt-text="Снимок экрана: лента &quot;Вставка&quot; в Excel в Интернете с выделенной кнопкой &quot;Мои надстройки&quot;.":::

1. Выберите пункт **Управление моими надстройками**, а затем выберите **Отправить мою надстройку**.

1. Выберите **Обзор...** и откройте корневой каталог проекта, созданный генератором Yeoman.

1. Выберите файл **manifest.xml** и нажмите **Открыть**, затем нажмите кнопку **Отправить**.

1. Теперь давайте оценим, как работает новая функция. В ячейке **B1** введите текст **=CONTOSO.GETSTARCOUNT("OfficeDev", "Excel-Custom-Functions")** и нажмите клавишу ВВОД. Результат в ячейке **B1** — это текущее количество звезд, отданных репозиторию [Excel-Custom-Functions Github](https://github.com/OfficeDev/Excel-Custom-Functions).

---