---
title: Создание области задач Excel с помощью Angular
description: ''
ms.date: 05/02/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 6c7baf237c4ca0be0e11428dcb2661b7b86ed39e
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771844"
---
# <a name="build-an-excel-task-pane-add-in-using-angular"></a>Создание области задач Excel с помощью Angular

Из этой статьи вы узнаете, как создать надстройку области Excel, используя Angular и API JavaScript для Excel.

## <a name="prerequisites"></a>Необходимые условия

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a>Создание проекта надстройки

Создайте проект надстройки Excel помощью генератора Yeoman. Выполните приведенную ниже команду и ответьте на вопросы, как показано ниже.

```command&nbsp;line
yo office
```

- **Выберите тип проекта:** `Office Add-in Task Pane project using Angular framework`
- **Выберите тип сценария:** `TypeScript`
- **Как вы хотите назвать надстройку?** `My Office Add-in`
- **Какое клиентское приложение Office должно поддерживаться?** `Excel`

![Генератор Yeoman](../images/yo-office-excel-angular-2.png)

После завершения работы мастера генератор создает проект и устанавливает вспомогательные компоненты Node.

## <a name="explore-the-project"></a>Знакомство с проектом

Проект надстройки, который вы создали с помощью генератора Yeoman, содержит образец кода для простейшей надстройки области задач. Если вы хотите ознакомиться с ключевыми компонентами проекта надстройки, откройте проект в редакторе кода и просмотрите файлы, перечисленные ниже. Когда вы будете готовы попробовать собственную надстройку, перейдите к следующему разделу.

- Файл **manifest.xml** в корневом каталоге проекта определяет настройки и возможности надстройки.
- Файл **./src/taskpane/app/app.component.html** содержит разметку HTML для области задач.
- Файл **./src/taskpane/taskpane.css** содержит код CSS, который применяется к содержимому области задач.
- Файл **./src/taskpane/app/app.component.ts** содержит код API JavaScript для Office, который упрощает взаимодействие между областью задач и Excel.

## <a name="try-it-out"></a>Проверка

1. Перейдите к корневой папке проекта.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)] 

3. В Excel выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.

    ![Кнопка надстройки Excel](../images/excel-quickstart-addin-3b.png)

4. Выберите любой диапазон ячеек на листе.

5. Внизу области задач выберите ссылку **Выполнить**, чтобы задать выбранному диапазону желтый цвет.

    ![Надстройка Excel](../images/excel-quickstart-addin-3c.png)

## <a name="next-steps"></a>Дальнейшие действия

Поздравляем! Вы успешно создали надстройку области задач Excel с помощью Angular! Чтобы узнать больше о возможностях надстроек Excel и создать более сложную надстройку, воспользуйтесь руководством по надстройкам Excel.

> [!div class="nextstepaction"]
> [Руководство по надстройкам Excel](../tutorials/excel-tutorial.md)

## <a name="see-also"></a>См. также

* [Руководство по надстройкам Excel](../tutorials/excel-tutorial-create-table.md)
* [Основные концепции программирования с помощью API JavaScript для Excel](../excel/excel-add-ins-core-concepts.md)
* [Примеры кода надстроек Excel](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [Справочник по API JavaScript для Excel](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
