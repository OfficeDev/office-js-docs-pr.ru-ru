---
title: Создание надстройки области задач Excel с помощью Angular
description: Узнайте, как создать простую надстройку области задач Excel, используя API JS для Office и Angular.
ms.date: 06/10/2022
ms.prod: excel
ms.localizationpriority: high
ms.openlocfilehash: 372649188d8f617f65e0c2eddc4d758047b1a2cc
ms.sourcegitcommit: 4f19f645c6c1e85b16014a342e5058989fe9a3d2
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/15/2022
ms.locfileid: "66091155"
---
# <a name="use-angular-to-build-an-excel-task-pane-add-in"></a>Создание надстройки области задач Excel с помощью Angular

Из этой статьи вы узнаете, как создать надстройку области Excel, используя Angular и API JavaScript для Excel.

## <a name="prerequisites"></a>Необходимые условия

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a>Создание проекта надстройки

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Выберите тип проекта:** `Office Add-in Task Pane project using Angular framework`
- **Выберите тип сценария:** `TypeScript`
- **Как вы хотите назвать надстройку?** `My Office Add-in`
- **Какое клиентское приложение Office должно поддерживаться?** `Excel`

![Снимок экрана: интерфейс командной строки генератора Yeoman надстроек Office, где в качестве типа проекта установлена инфраструктура Angular.](../images/yo-office-excel-angular-2.png)

После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a>Знакомство с проектом

Проект надстройки, который вы создали с помощью генератора Yeoman, содержит образец кода для простой надстройки области задач. Если вы хотите ознакомиться с ключевыми компонентами проекта надстройки, откройте проект в редакторе кода и просмотрите файлы, перечисленные ниже. Когда вы будете готовы попробовать собственную надстройку, перейдите к следующему разделу.

- Файл **manifest.xml** в корневом каталоге проекта определяет настройки и возможности надстройки. Дополнительные сведения о файле **manifest.xml** см. в статье [XML-манифест надстроек Office](../develop/add-in-manifests.md).
- Файл **./src/taskpane/app/app.component.html** содержит разметку HTML для области задач.
- Файл **./src/taskpane/taskpane.css** содержит код CSS, который применяется к содержимому области задач.
- Файл **./src/taskpane/app/app.component.ts** содержит код API JavaScript для Office, который упрощает взаимодействие между областью задач и Excel.

## <a name="try-it-out"></a>Проверка

1. Перейдите к корневой папке проекта.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

1. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)]

1. В Excel выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.

    ![Снимок экрана: меню "Главная" в Excel с выделенной кнопкой "Показать область задач".](../images/excel-quickstart-addin-3b.png)

1. Выберите любой диапазон ячеек на листе.

1. Внизу области задач выберите ссылку **Выполнить**, чтобы задать выбранному диапазону желтый цвет.

    ![Снимок экрана: Excel с открытой областью задач надстройки и выделенной кнопкой "Выполнить".](../images/excel-quickstart-addin-3c.png)

## <a name="next-steps"></a>Дальнейшие действия

Поздравляем, вы успешно создали надстройку панели задач Excel с помощью Angular! Чтобы узнать больше о возможностях надстроек Excel и создать более сложную надстройку, воспользуйтесь руководством по надстройкам Excel.

> [!div class="nextstepaction"]
> [Руководство по надстройкам Excel](../tutorials/excel-tutorial.md)

## <a name="see-also"></a>См. также

- [Обзор платформы надстроек Office](../overview/office-add-ins.md)
- [Разработка надстроек Office](../develop/develop-overview.md)
- [Объектная модель JavaScript для Excel в надстройках Office](../excel/excel-add-ins-core-concepts.md)
- [Примеры кода надстроек Excel](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
- [Справочник по API JavaScript для Excel](../reference/overview/excel-add-ins-reference-overview.md)
- [Использование Visual Studio Code для публикации](../publish/publish-add-in-vs-code.md#using-visual-studio-code-to-publish)
