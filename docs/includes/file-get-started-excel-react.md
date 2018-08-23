# <a name="build-an-excel-add-in-using-react"></a>Создание надстройки Excel с помощью React

В этой статье описывается процесс создания надстройки Excel с помощью React и API JavaScript для Excel.

## <a name="environment"></a>Среда

- **Классическое приложение Office.** Убедитесь, что у вас установлена ​​последняя версия Office. Команды надстроек требуют сборку 16.0.6769.0000 или более позднюю (рекомендуется сборка **16.0.6868.0000**). Узнайте, как [установить последнюю версию приложений Office](http://aka.ms/latestoffice). 
 
- **Office Online.** Не требуется выполнять дополнительную настройку. Обратите внимание, что поддержка команд в Office Online для рабочих и учебных учетных записей предоставляется в тестовом режиме.

## <a name="prerequisites"></a>Необходимые компоненты

- [Node.js](https://nodejs.org)

- Глобально установите последнюю версию [Yeoman](https://github.com/yeoman/yo) и [генератор Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office).
    ```bash
    npm install -g yo generator-office
    ```

### <a name="create-the-web-app"></a>Создание веб-приложения

1. Создайте на локальном диске папку и назовите ее **my-addin**. В ней вы будете создавать файлы для приложения.

2. Перейдите к папке приложения.

    ```bash
    cd my-addin
    ```

3. Используя генератор Yeoman, создайте файл манифеста для надстройки. Выполните приведенную ниже команду и ответьте на вопросы, как показано на следующем снимке экрана.

    ```bash
    yo office
    ```

    - **Выберите тип проекта:** `Office Add-in project using React framework`
    - **Как вы хотите назвать надстройку?:** `My Office Add-in`
    - **Какое клиентское приложение Office должно поддерживаться?:** `Excel`

    ![Генератор Yeoman](../images/yo-office-excel-react.png)
    
    После завершения работы мастера генератор создаст проект и установит поддерживающие компоненты узла.

4.  Откройте **src/components/App.tsx**, найдите комментарий "Обновить цвет заливки", а затем измените цвет заливки с 'желтого' на 'синий' и сохраните файл. 

    ```js
    range.format.fill.color = 'blue'

    ```

5. В блоке `return` функции `render` внутри **src/components/App.tsx** обновите `<Herolist>` в соответствии с приведенным ниже кодом и сохраните файл. 

    ```js
      <HeroList message='Discover what My Office Add-in can do for you today!' items={this.state.listItems}>
        <p className='ms-font-l'>Choose the button below to set the color of the selected range to blue. <b>Set color</b>.</p>
        <Button className='ms-welcome__action' buttonType={ButtonType.hero} iconProps={{ iconName: 'ChevronRight' }} onClick={this.click}>Run</Button>
    </HeroList>
    ```

6. Сделайте так, чтобы операционная система компьютера разработки доверяла сертификату. Для этого выполните действия, описанные в статье [Добавление самозаверяющих сертификатов в качестве доверенного корневого сертифтката](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).

7. Загрузите неопубликованную надстройку, чтобы она отобразилась в Excel. В терминале выполните следующую команду: 
    
    ```bash
    npm run sideload
    ```

## <a name="try-it-out"></a>Проверка

1. Выполните в терминале приведенную ниже команду, чтобы запустить сервер разработки.

    Windows:
    ```bash
    npm start
    ```

2. В Excel выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.

    ![Кнопка надстройки Excel](../images/excel-quickstart-addin-2b.png)

3. Выберите любой диапазон ячеек на листе.

4. В области задач нажмите кнопку **Выбрать цвет**, чтобы сделать выбранный диапазон синим.

    ![Надстройка Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a>Дальнейшие действия

Поздравляем, вы успешно создали надстройку Excel с помощью React! Чтобы узнать больше о возможностях надстроек Excel и создать более сложную надстройку, воспользуйтесь руководством по надстройкам Excel.

> [!div class="nextstepaction"]
> [Руководство по надстройкам Excel](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a>См. также

* [Руководство по надстройкам Excel](../tutorials/excel-tutorial-create-table.md)
* [Основные понятия API JavaScript для Excel](../excel/excel-add-ins-core-concepts.md)
* [Примеры кода надстроек Excel](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [Справочник по API JavaScript для Excel](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
