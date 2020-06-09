---
title: Создание области задач Excel с помощью Vue
description: Узнайте, как создать простую надстройку области задач Excel, используя API JS для Office и Vue.
ms.date: 04/14/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 49c38a8b9055692f40fed9dc110528694e3fbf55
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608924"
---
# <a name="build-an-excel-task-pane-add-in-using-vue"></a>Создание области задач Excel с помощью Vue

Из этой статьи вы узнаете, как создать надстройку области Excel, используя Vue и API JavaScript для Excel.

## <a name="prerequisites"></a>Необходимые условия

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- Установите [Vue CLI](https://cli.vuejs.org/) глобально.

  ```command&nbsp;line
  npm install -g @vue/cli
  ```

## <a name="generate-a-new-vue-app"></a>Создание нового приложения Vue

Используйте Vue CLI, чтобы создать новое приложение Vue. С помощью терминала выполните следующую команду.

```command&nbsp;line
vue create my-add-in
```

Затем выберите параметр `default`. Если в качестве пакета предлагается использовать Yarn или NPM, можно выбрать любой из этих вариантов.

## <a name="generate-the-manifest-file"></a>Создание файла манифеста

У каждой надстройки должен быть файл манифеста, в нем определяются ее параметры и возможности.

1. Перейдите к папке приложения.

    ```command&nbsp;line
    cd my-add-in
    ```

2. С помощью генератора Yeoman создайте файл манифеста для надстройки, выполнив следующую команду:

    ```command&nbsp;line
    yo office
    ```

    > [!NOTE]
    > При выполнении команды `yo office` может появиться запрос о политиках сбора данных генератора Yeoman и средств CLI надстройки Office. Используйте предоставленные сведения, чтобы ответить на запросы подходящим образом. Если в ответ на второй запрос выбран вариант **Выход**, потребуется снова выполнить команду `yo office`, когда вы будете готовы создать проект надстройки.

    При появлении запроса предоставьте следующую информацию для создания проекта надстройки:

    - **Выберите тип проекта:** `Office Add-in project containing the manifest only`
    - **Как вы хотите назвать надстройку?** `My Office Add-in`
    - **Какое клиентское приложение Office должно поддерживаться?** `Excel`

    ![Генератор Yeoman](../images/yo-office-manifest-only-vue.png)

После завершения работы мастера создается папка `My Office Add-in`, содержащая файл `manifest.xml`. В конце краткого руководства вам потребуется использовать манифест для загрузки без публикации и тестирования вашей надстройки.

> [!TIP]
> Вы можете игнорировать инструкции по *дальнейшим действиям*, предоставляемые генератором Yeoman после создания проекта надстройки. Пошаговые инструкции этой статьи содержат все сведения, необходимые для завершения этого учебного курса.

## <a name="secure-the-app"></a>Защита приложения

[!include[HTTPS guidance](../includes/https-guidance.md)]

1. Чтобы включить HTTPS для своего приложения, создайте файл `vue.config.js` в корневой папке проекта Vue со следующим содержанием:

    ```js
    var fs = require("fs");
    var path = require("path");
    var homedir = require('os').homedir()
  
    module.exports = {
      devServer: {
        port: 3000,
        https: true,
        key: fs.readFileSync(path.resolve(`${homedir}/.office-addin-dev-certs/localhost.key`)),
        cert: fs.readFileSync(path.resolve(`${homedir}/.office-addin-dev-certs/localhost.crt`)),
        ca: fs.readFileSync(path.resolve(`${homedir}/.office-addin-dev-certs/ca.crt`))
      }
    }
    ```

2. В терминале выполните следующую команду, чтобы установить сертификаты надстройки.

   ```command&nbsp;line
   npx office-addin-dev-certs install
   ```

## <a name="update-the-app"></a>Обновление приложения

1. Откройте файл `public/index.html` и добавьте следующий тег `<script>` непосредственно перед тегом `</head>`:

   ```html
   <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
   ```

2. Откройте файл `src/main.js` и замените его содержимое указанным ниже кодом:

   ```js
   import Vue from 'vue';
   import App from './App.vue';

   Vue.config.productionTip = false;

   window.Office.initialize = () => {
     new Vue({
       render: h => h(App)
     }).$mount('#app');
   };
   ```

3. Откройте файл `src/App.vue` и замените содержимое файла указанным ниже кодом:

   ```html
   <template>
     <div id="app">
       <div class="content">
         <div class="content-header">
           <div class="padding">
             <h1>Welcome</h1>
           </div>
         </div>
         <div id="content-main">
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

1. Используя терминал, выполните приведенную ниже команду, чтобы запустить сервер разработки.

   ```command&nbsp;line
   npm run serve
   ```

2. В веб-браузере перейдите по адресу `https://localhost:3000` (обратите внимание на `https`). Если страница `https://localhost:3000` пуста и не содержит ошибок сертификата, значит, она работает. Приложение Vue подключается после запуска Office, поэтому в нем отображаются только элементы из среды Excel.

## <a name="try-it-out"></a>Проверка

1. Следуя указаниям для нужной платформы, загрузите неопубликованную надстройку в Excel.

   - [Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
   - Веб-браузер: [загрузка неопубликованных надстроек Office в Office в Интернете](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)
   - [iPad и Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

2. В Excel выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.

   ![Кнопка надстройки Excel](../images/excel-quickstart-addin-2a.png)

3. Выберите любой диапазон ячеек на листе.

4. В области задач нажмите кнопку **Set color** (Задать цвет), чтобы сделать выбранный диапазон зеленым.

   ![Надстройка Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a>Дальнейшие действия

Поздравляем! Вы успешно создали надстройку области задач Excel с помощью Vue! Чтобы узнать больше о возможностях надстроек Excel и создать более сложную надстройку, воспользуйтесь руководством по надстройкам Excel.

> [!div class="nextstepaction"]
> [Руководство по надстройкам Excel](../tutorials/excel-tutorial.md)

## <a name="see-also"></a>См. также

* [Обзор платформы надстроек Office](../overview/office-add-ins.md)
* [Создание надстроек Office](../overview/office-add-ins-fundamentals.md)
* [Разработка надстроек Office](../develop/develop-overview.md)
* [Основные концепции программирования с помощью API JavaScript для Excel](../excel/excel-add-ins-core-concepts.md)
* [Примеры кода надстроек Excel](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [Справочник по API JavaScript для Excel](../reference/overview/excel-add-ins-reference-overview.md)
