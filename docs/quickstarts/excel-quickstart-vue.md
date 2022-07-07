---
title: Создание надстройки области задач Excel с помощью Vue
description: Узнайте, как создать простую надстройку области задач Excel, используя API JS для Office и Vue.
ms.date: 06/10/2022
ms.prod: excel
ms.localizationpriority: high
ms.openlocfilehash: 8fb4bd545e1fab44884dd4a5dc388910d71c8336
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/06/2022
ms.locfileid: "66659796"
---
# <a name="use-vue-to-build-an-excel-task-pane-add-in"></a>Создание надстройки области задач Excel с помощью Vue

Из этой статьи вы узнаете, как создать надстройку области Excel, используя Vue и API JavaScript для Excel.

## <a name="prerequisites"></a>Необходимые условия

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- Установите [Vue CLI](https://cli.vuejs.org/) глобально. В терминале выполните следующую команду:

  ```command&nbsp;line
  npm install -g @vue/cli
  ```

## <a name="generate-a-new-vue-app"></a>Создание нового приложения Vue

Используйте интерфейс командной строки Vue, чтобы создать новое приложение Vue.

```command&nbsp;line
vue create my-add-in
```

Затем выберите предустановку `Default` для "Vue 3" (также можно использовать "Vue 2").

## <a name="generate-the-manifest-file"></a>Создание файла манифеста

У каждой надстройки должен быть файл манифеста, в нем определяются ее параметры и возможности.

1. Перейдите к папке приложения.

    ```command&nbsp;line
    cd my-add-in
    ```

1. Используя генератор Yeoman, создайте файл манифеста для надстройки.

    ```command&nbsp;line
    yo office
    ```

    > [!NOTE]
    > При выполнении команды `yo office` могут появиться запросы о политиках сбора данных генератора Yeoman и средств командной строки надстройки Office. Используйте предоставленные сведения, чтобы ответить на запросы подходящим образом. Если в ответ на второй запрос выбран вариант **Выход**, потребуется снова выполнить команду `yo office`, когда вы будете готовы создать проект надстройки.

    При появлении запроса предоставьте следующую информацию для создания проекта надстройки.

    - **Выберите тип проекта:** `Office Add-in project containing the manifest only`
    - **Как вы хотите назвать надстройку?** `My Office Add-in`
    - **Какое клиентское приложение Office должно поддерживаться?** `Excel`

    ![Снимок экрана: интерфейс командной строки генератора Yeoman надстроек Office, где в качестве типа проекта установлен только манифест.](../images/yo-office-manifest-only-vue.png)

По завершении мастер создает папку **My Office Add-in**, содержащую файл **manifest.xml**. Вы воспользуетесь манифестом для загрузки вашей неопубликованной надстройки и для ее тестирования.

> [!TIP]
> Вы можете игнорировать инструкции по *дальнейшим действиям*, предоставляемые генератором Yeoman после создания проекта надстройки. Пошаговые инструкции этой статьи содержат все сведения, необходимые для завершения этого учебного курса.

## <a name="secure-the-app"></a>Защита приложения

[!include[HTTPS guidance](../includes/https-guidance.md)]

1. Включите протокол HTTPS для вашего приложения. В корневой папке проекта Vue создайте файл **vue.config.js** со следующим содержимым.

    ```js
    var fs = require("fs");
    var path = require("path");
    var homedir = require('os').homedir()
  
    module.exports = {
      devServer: {
        port: 3000,
        https: {
          key: fs.readFileSync(path.resolve(`${homedir}/.office-addin-dev-certs/localhost.key`)),
          cert: fs.readFileSync(path.resolve(`${homedir}/.office-addin-dev-certs/localhost.crt`)),
          ca: fs.readFileSync(path.resolve(`${homedir}/.office-addin-dev-certs/ca.crt`)),
         }
       }
    }
    ```

1. Установите сертификаты надстройки.

   ```command&nbsp;line
   npx office-addin-dev-certs install
   ```

## <a name="explore-the-project"></a>Знакомство с проектом

Проект надстройки, который вы создали с помощью генератора Yeoman, содержит образец кода для простой надстройки области задач. Если вы хотите ознакомиться с ключевыми компонентами проекта надстройки, откройте проект в редакторе кода и просмотрите файлы, перечисленные ниже. Когда вы будете готовы попробовать собственную надстройку, перейдите к следующему разделу.

- Файл **manifest.xml** в корневом каталоге проекта определяет настройки и возможности надстройки. Дополнительные сведения о файле **manifest.xml** см. в статье [XML-манифест надстроек Office](../develop/add-in-manifests.md).
- Файл **./src/App.vue** содержит разметку HTML для области задач, таблицу стилей CSS, применяемую к содержимому области задач и код API JavaScript Office, обеспечивающий взаимодействие между областью задач и Excel.

## <a name="update-the-app"></a>Обновите приложение

1. Откройте файл **./public/index.html** и добавьте следующий тег `<script>` непосредственно перед тегом `</head>`.

   ```html
   <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
   ```

1. Откройте файл **manifest.xml** и найдите теги `<bt:Urls>` внутри тега **\<Resources\>**. Найдите тег `<bt:Url>` с идентификатором `Taskpane.Url` и обновите его атрибут `DefaultValue`. Новое значение `DefaultValue` — `https://localhost:3000/index.html`. Весь обновленный тег должен совпадать со следующей строкой.

   ```html
   <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/index.html" />
   ```

1. Откройте файл **src/main.js** и замените его содержимое на следующий код.

   ```js
   import { createApp } from 'vue'
   import App from './App.vue'

   window.Office.onReady(() => {
       createApp(App).mount('#app');
   });
   ```

1. Откройте файл **./src/App.vue** и замените его содержимое на следующий код.

   ```html
   <template>
     <div id="app">
       <div class="content">
         <div class="content-header">
           <div class="padding">
             <h1>Welcome</h1>
           </div>
         </div>
         <div class="content-main">
           <div class="padding">
             <p>
               Choose the button below to set the color of the selected range to
               green.
             </p>
             <br />
             <h3>Try it out</h3>
             <button @click="onSetColor">Set color</button>
           </div>
         </div>
       </div>
     </div>
   </template>

   <script>
     export default {
       name: 'App',
       methods: {
         onSetColor() {
           window.Excel.run(async context => {
             const range = context.workbook.getSelectedRange();
             range.format.fill.color = 'green';
             await context.sync();
           });
         }
       }
     };
   </script>

   <style>
     .content-header {
       background: #2a8dd4;
       color: #fff;
       position: absolute;
       top: 0;
       left: 0;
       width: 100%;
       height: 80px;
       overflow: hidden;
     }

     .content-main {
       background: #fff;
       position: fixed;
       top: 80px;
       left: 0;
       right: 0;
       bottom: 0;
       overflow: auto;
     }

     .padding {
       padding: 15px;
     }
   </style>
   ```

## <a name="start-the-dev-server"></a>Запуск сервера разработки

1. Запустите сервер разработки

   ```command&nbsp;line
   npm run serve
   ```

1. В веб-браузере перейдите по адресу `https://localhost:3000` (обратите внимание на `https`). Если страница `https://localhost:3000` пуста, а ошибки сертификата отсутствуют, значит, эта страница работает. Приложение Vue подключается после запуска Office, поэтому в нем отображаются только элементы из среды Excel.

## <a name="try-it-out"></a>Проверка

1. Запустите надстройку и загрузите неопубликованную надстройку в Excel. Следуйте инструкциям для используемой вами платформы:

   - [Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
   - Веб-браузер: [загрузка неопубликованных надстроек Office в Office в Интернете](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)
   - [iPad и Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

1. Откройте область задач надстройки в Excel. На вкладке **Главная** нажмите кнопку **Показать область задач**.

   ![Снимок экрана: меню "Главная" в Excel с выделенной кнопкой "Показать область задач".](../images/excel-quickstart-addin-2a.png)

1. Выберите любой диапазон ячеек на листе.

1. Установите зеленый цвет для выбранного диапазона. В области задач надстройки нажмите кнопку **Задать цвет**.

   ![Снимок экрана: Excel с открытой областью задач надстройки.](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a>Дальнейшие действия

Поздравляем, вы успешно создали надстройку панели задач Excel с помощью Vue! Чтобы узнать больше о возможностях надстроек Excel и создать более сложную надстройку, воспользуйтесь руководством по надстройкам Excel.

> [!div class="nextstepaction"]
> [Руководство по надстройкам Excel](../tutorials/excel-tutorial.md)

## <a name="see-also"></a>См. также

- [Обзор платформы надстроек Office](../overview/office-add-ins.md)
- [Разработка надстроек Office](../develop/develop-overview.md)
- [Объектная модель JavaScript для Excel в надстройках Office](../excel/excel-add-ins-core-concepts.md)
- [Примеры кода надстроек Excel](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
- [Справочник по API JavaScript для Excel](../reference/overview/excel-add-ins-reference-overview.md)
- [Использование Visual Studio Code для публикации](../publish/publish-add-in-vs-code.md#using-visual-studio-code-to-publish)