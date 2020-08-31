---
ms.date: 08/25/2020
title: Настройка надстройки Excel для совместного использования среды выполнения браузера
ms.prod: excel
description: Настройте надстройку Excel, чтобы предоставить общий доступ к среде выполнения браузера и запускать код ленты, области задач и пользовательских функций в одной и той же среде выполнения.
localization_priority: Priority
ms.openlocfilehash: 08e4155b7f79101f8a61b323c623b5cb6b86decf
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292638"
---
# <a name="configure-your-excel-add-in-to-use-a-shared-javascript-runtime"></a>Настройка надстройки Excel для использования общей среды выполнения JavaScript

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

При запуске Excel на компьютере с Windows или на Mac надстройка запустит код для кнопок ленты, пользовательских функций и области задач в отдельных средах выполнения JavaScript. Из-за этого возникают ограничения, например невозможность удобно предоставлять общий доступ к глобальным данным и отсутствие доступа ко всей функциональности CORS для пользовательской функции.

Но вы можете настроить вашу надстройку Excel, предоставив общий доступ к коду в общей среде выполнения  JavaScript. Это позволяет повысить слаженность работы всей вашей надстройки и обеспечить доступ к DOM и CORS из всех ее частей. Кроме того, это позволяет запускать код при открытии документа и после закрытия области задач. Чтобы настроить надстройку для использования общей среды выполнения, следуйте инструкциям, приведенным в этой статье.

## <a name="create-the-add-in-project"></a>Создание проекта надстройки

Если вы начинаете новый проект, выполните указанные ниже действия, чтобы с помощью генератора Yeoman создать проект надстройки Excel. Выполните приведенную ниже команду и ответьте на вопросы, как показано ниже.

```command line
yo office
```

- Выберите тип проекта: **проект надстройки пользовательских функций Excel**
- Выберите тип сценария: **JavaScript**
- Как вы хотите назвать надстройку? **Моя надстройка Office**

![Снимок экрана: ответы на вопросы Office о создании проекта надстройки.](../images/yo-office-excel-project.png)

После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.

## <a name="configure-the-manifest"></a>Настройка манифеста

Выполните указанные ниже действия для нового или существующего проекта, чтобы настроить его для использования общей среды выполнения.

1. Запустите Visual Studio Code и откройте проект **Моя надстройка Office**.
2. Откройте файл **manifest.xml**.
3. Найдите раздел `<VersionOverrides>` и добавьте следующий раздел `<Runtimes>`. Время существования должно быть **длительным**, чтобы пользовательские функции могли работать даже после закрытия области задач. Атрибут resid равен `ContosoAddin.Url` и ссылается на строку в разделе ресурсов далее. Можно использовать любое значение resid, но оно должно соответствовать resid других элементов в элементах вашей надстройки.

   ```xml
   <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
     <Hosts>
       <Host xsi:type="Workbook">
       <Runtimes>
         <Runtime resid="ContosoAddin.Url" lifetime="long" />
       </Runtimes>
       <AllFormFactors>
   ```

4. В элементе `<Page>` замените расположение источника с **Functions.Page.Url** на **ContosoAddin.Url**. Этот resid соответствует элементу resid `<Runtime>`. Обратите внимание: если у вас нет пользовательских функций, то у вас не будет элемента **Page**, и этот шаг можно пропустить.

   ```xml
   <AllFormFactors>
   ...
   <Page>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Page>
   ...
   ```

5. В разделе `<DesktopFormFactor>` замените **FunctionFile** с **Commands.Url** на **ContosoAddin.Url**. Обратите внимание: если у вас нет команд действий, то у вас не будет элемента **FunctionFile**, и этот шаг можно пропустить.

   ```xml
   <DesktopFormFactor>
   <GetStarted>
   ...
   </GetStarted>
   <FunctionFile resid="ContosoAddin.Url"/>
   ```

6. В разделе `<Action>` измените расположение источника с **Taskpane.Url** на **ContosoAddin.Url**. Обратите внимание: если у вас нет области задач, то у вас не будет действия **ShowTaskpane**, и этот шаг можно пропустить.

   ```xml
   <Action xsi:type="ShowTaskpane">
   <TaskpaneId>ButtonId1</TaskpaneId>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Action>
   ```

7. Добавьте новый **Url-идентификатор** для **ContosoAddin.Url**, указывающий на **taskpane.html**.

   ```xml
   <bt:Urls>
   <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
   ...
   <bt:Url id="ContosoAddin.Url" DefaultValue="https://localhost:3000/dist/taskpane.html"/>
   ...
   ```

8. Убедитесь в том, что в файле taskpane.html есть тег `<script>`, ссылающийся на файл dist/functions.js. Ниже приведен пример.

   ```html
   <script type="text/javascript" src="/dist/functions.js" ></script>
   ```

   > [!NOTE]
   > Если для вставки тегов сценариев надстройка использует Webpack и HtmlWebpackPlugin, как это делают надстройки, созданные генератором Yeoman (см. раздел [Создание проекта надстройки](#create-the-add-in-project) выше), то вам необходимо обеспечить включение модуля functions.js в массив `chunks`, как в следующем примере.
   >
   > ```javascript
   > new HtmlWebpackPlugin({
   >     filename: "taskpane.html",
   >     template: "./src/taskpane/taskpane.html",
   >     chunks: ["polyfill", "taskpane", “functions”]
   > }),
   >```

9. Сохраните изменения и перестройте проект.

   ```command line
   npm run build
   ```

## <a name="runtime-lifetime"></a>Срок существования среды выполнения

Добавляя элемент `Runtime`, вы также задаете срок существования со значением `long` или `short`. Установите значение `long`, чтобы воспользоваться такими функциями, как запуск надстройки при открытии документа, продолжение выполнения кода после закрытия области задач или использование CORS и DOM из пользовательских функций.

>[!NOTE]
> По умолчанию используется значение срока жизни `short`, но мы рекомендуем использовать `long` в надстройках Excel. Если вы настроите в этом примере для среды выполнения значение `short`, ваша надстройка Excel запустится при нажатии одной из кнопок на ленте, но может завершить работу после окончания функционирования обработчика ленты. Точно так же, надстройка запустится при открытии области задач, но может завершить работу после закрытия области задач.

```xml
<Runtimes>
  <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

>[!NOTE]
> Если в манифесте надстройки есть элемент `Runtimes` (требуемый для общей среды выполнения), она использует Internet Explorer 11 независимо от того, какая у вас версия Windows или Microsoft 365. Дополнительные сведения см. в статье [Runtimes](../reference/manifest/runtimes.md).

## <a name="multiple-task-panes"></a>Несколько областей задач

Не планируйте использовать в своей надстройке несколько областей задач, если предполагается использование общей среды выполнения. Общая среда выполнения поддерживает только одну область задач. Обратите внимание: любая область задач без `<TaskpaneID>` считается другой областью задач.

## <a name="next-steps"></a>Дальнейшие действия

- Подробные сведения об использовании API JavaScript для Excel и пользовательских функций Excel в общей среде выполнения см. в статье [Вызов API Excel из пользовательской функции](call-excel-apis-from-custom-function.md).
- Изучите пример PnP [Управление интерфейсом ленты и области задач, а также запуск кода при открытии документа](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-shared-runtime-scenario), чтобы ознакомиться с масштабным примером работы общей среды выполнения JavaScript.

## <a name="see-also"></a>См. также

- [Обзор: запуск кода надстройки в общей среде выполнения JavaScript](custom-functions-shared-overview.md)
