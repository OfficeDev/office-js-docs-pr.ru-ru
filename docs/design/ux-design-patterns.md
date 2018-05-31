---
title: Конструктивные шаблоны пользовательского интерфейса надстроек Office
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: c8ec23db5e7c4c571babff94bdc617b78340d965
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437628"
---
# <a name="ux-design-pattern-templates-for-office-add-ins"></a>Конструктивные шаблоны пользовательского интерфейса надстроек Office

[Проект конструктивных шаблонов пользовательского интерфейса для надстроек Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code "Проект конструктивных шаблонов пользовательского интерфейса для надстроек Office") включает файлы HTML, JavaScript и CSS, с помощью которых можно создать пользовательский интерфейс надстройки.   

С помощью проекта конструктивных шаблонов пользовательского интерфейса можно:

* применять решения в распространенных клиентских сценариях;
* следовать рекомендациям по оформлению;
* внедрять компоненты и стили [Office UI Fabric](https://dev.office.com/fabric#/get-started);
* создавать надстройки, внешний вид которых согласован со стандартным пользовательским интерфейсом Office.  

## <a name="using-the-ux-design-patterns"></a>Использование конструктивных UX шаблонов пользовательского интерфейса

Вы можете использовать [Набор инструментов надстроек Office Add-ins Design](https://aka.ms/addins_toolkit) вместе с набором инструментов надстроек[Fabric Design](https://aka.ms/fabric-toolkit) как руководство для  создания собственной надстройки Office. Вы также можете добавить [исходный код](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates) напрямую к вашему проекту.

Чтобы создать пользовательский интерфейс надстройки на основе спецификаций, выполните указанные ниже действия.

1. Скачайте файлы активов разработки и начните создание собственного пользовательского интерфейса:
    * [Набор инструментов надстроек Office Add-ins Design](https://aka.ms/addins_toolkit)
    * [Набор инструментов надстроек Fabric Design](https://aka.ms/fabric-toolkit)

2. Инструкции см. в следующих статьях:
    * Рекомендации по [проектированию надстроек Office](add-in-design.md)
    * [Наборы инструментов Office UI Fabric](https://developer.microsoft.com/en-us/fabric#/resources)

> [!NOTE]
> Некоторые  UX-шаблоны в наборе инструментов надстроек не соответствуют  UX-шаблонам проектирования, подробно описанным ниже. Мы планируем выпустить обновленную документацию, которая будет согласована с набором инструментальных средств.

Чтобы добавить исходный код, сделайте следующее:

1. Создайте копию [репозитория проекта конструктивных шаблонов пользовательского интерфейса для надстроек Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code "Проект конструктивных шаблонов пользовательского интерфейса для надстроек Office").
2. Скопируйте [папку с ресурсами](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/assets) и папку с кодом выбранного шаблона в проект надстройки.  
3. Внедрите шаблон в надстройку. Например:
    - Измените расположение исходного кода или URL-адрес команд надстройки в манифесте.
    - Используйте шаблон UX-дизайна в качестве шаблона для других страниц.
    - Создайте ссылки на конструктивный шаблон пользовательского интерфейса или из него.

> [!NOTE]
> Некоторые спецификации шаблонов пользовательского интерфейса отличаются от исходного кода. Мы работаем над согласованием всех активов. Кроме того, обратите внимание, что некоторые спецификации представлены как архивные. Мы оцениваем значение этих архивных спецификаций для платформы. Каждый шаблон уникален и представляет шаблон взаимодействия. Шаблоны не должны перекрывать друг друга. Они также должны отличаться от компонентов Office UI Fabric.


## <a name="types-of-ux-design-patterns"></a>Типы конструктивных шаблонов пользовательского интерфейса
### <a name="generic-pages"></a>Универсальные страницы

Шаблоны универсальных страниц можно применить к любой странице надстройки. У них нет особого назначения. Примером страницы особого назначения может служить любой из шаблонов первого запуска. В следующем списке описаны доступные универсальные страницы:

* **Целевая страница** — это стандартная страница надстройки, например страница, которая открывается после первого запуска или входа. 
    * Рекомендации по внедрению [языка дизайна Office](add-in-design-language.md) в надстройке.
    * [Код целевой страницы](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/generic/landing-page)
* **Панель с фирменной символикой** — целевая страница с изображением, которое находится в нижнем колонтитуле и представляет торговую марку. 
    * [Спецификация панели с фирменной символикой](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/brand-bar.md)
    * [Код панели с фирменной символикой](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/generic/brand-bar)

<table>
 <tr><th>Целевая</th><th>Панель с фирменной символикой</th></tr>
 <tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/generic/landing-page"><img src="../images/landing-pages.png" alt="landing page" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/generic/brand-bar"><img src="../images/word-brand-bar.png" alt="brand bar" style="width: 264px;"/></A></td></tr>
 </table>
 
### <a name="first-run-experience"></a>Интерфейс при первом запуске

Интерфейс при первом запуске — это интерфейс, отображаемый для пользователя, когда тот запускает надстройку в первый раз. Доступны следующие конструктивные шаблоны для первого запуска: 

* **Действия, необходимые для запуска.** Предоставляет пользователям упорядоченный список действий, которые необходимо выполнить, чтобы начать использовать надстройку. 
    * [Спецификация действий для запуска](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/fre_stepsToStart.pdf) (этот конструктивный шаблон заархивирован. Пока мы оцениваем его ценность, см. [спецификацию значений при первом запуске](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/value-placemat.md).)  
    * [Код действий, необходимых для запуска](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/instruction-step)
* **Решаемые задачи.** Разъясняет, какие задачи можно решить с помощью надстройки.
    * [Спецификация значения](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/value-placemat.md)
    * [Код решаемых задач](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/value-placemat)
* **Видео.** Показывает пользователям видеоролик перед тем, как они начнут использовать вашу надстройку.
    * [Спецификация видео](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/video-placemat.md)
    * [Код видео](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/video-placemat)
* **Пошаговое руководство.** Рассказывает пользователям о ряде функций или предоставляет определенные сведения, прежде чем они начнут использовать надстройку.
    * [Спецификация карусели](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/carousel.md) (обратите внимание, что этот конструктивный шаблон пользовательского интерфейса переименован в "Карусель". Ранее в спецификациях он носил имя "Панель разбиения на страницы". В ресурсах с кодом он упоминается под названием "Пошаговое руководство по первому запуску". 
    * [Код пошагового руководства](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/walkthrough)

В [AppSource](https://docs.microsoft.com/en-us/office/dev/store/use-the-seller-dashboard-to-submit-to-the-office-store) предусмотрена система управления пробными версиями надстроек, но если вам нужно управлять пользовательским интерфейсом такой пробной версии, используйте следующие шаблоны:

* **Пробная версия.** Показывает пользователям, как начать работу с пробной версией надстройки.
    * [Спецификация пробной версии](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/fre_trialVersion.pdf) (этот конструктивный шаблон заархивирован. Пока мы оцениваем его ценность, руководствуйтесь данным PDF-файлом.)
    * [Код пробной версии](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/trial-placemat)
* **Пробная функция.** Сообщает пользователю, что функция, которую он пытается использовать, недоступна в пробной версии надстройки. Кроме того, этот шаблон можно использовать, если надстройка предоставляется бесплатно, но в ней есть функция, для использования которой требуется подписка. С помощью этого шаблона вы также можете предоставлять ограниченную версию после окончания пробного периода.
    * [Спецификация функций пробной версии](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/fre_trialFeature.pdf) (этот конструктивный шаблон заархивирован. Пока мы оцениваем его ценность, руководствуйтесь данным PDF-файлом.)
    * [Код функций пробной версии](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/trial-placemat-feature)

> [!IMPORTANT]
> Если вы решите самостоятельно управлять подпиской, не применяя для этого AppSource, добавьте тег **Может потребоваться дополнительная покупка** в примечания по тестированию на Панели мониторинга продаж.

Решите, как часто необходимо применять интерфейс, используемый при первом запуске: один раз или несколько. Например, если пользователи используют надстройку периодически, они могут забыть, как ею пользоваться, поэтому интерфейс первого запуска может пригодиться более одного раза. 

 <table>
 <tr><th>Действия, необходимые для запуска</th><th>Решаемые задачи</th><th>Видео</th></tr>
 <tr>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/instruction-step"><img src="../images/instruction-steps.png" alt="instruction steps" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/value-placemat"><img src="../images/value-placemats.png" alt="value placemat" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/video-placemat"><img src="../images/video-placemats.png" alt="video placemat" style="width: 250px;"/></A></td></tr>
 </table>

 <table>
 <tr><th>Первая страница пошагового руководства</th><th>Пробная версия</th><th>Пробная функция</th></tr>
 <tr>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/walkthrough"><img src="../images/walkthrough01.png" alt="walkthrough 1" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/trial-placemat"><img src="../images/trial-placemats.png" alt="trial placemat" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/trial-placemat-feature"><img src="../images/trial-placemats-feature.png" alt="trial placemat feature" style="width: 250px;"/></A></td></tr>
 </table> 

### <a name="navigation"></a>Навигация

Пользователям нужно переключаться между разными страницами надстройки. В приведенных ниже шаблонах навигации показаны различные варианты упорядочения страниц и команд надстройки.

* **Страница с кнопками "Назад" и "Вперед".** Показывает область задач с кнопками "Назад" и "Вперед". Этот шаблон поможет гарантировать, что пользователи выполняют действия по порядку.
    * [Спецификация страницы с кнопками "Назад" и "Вперед"](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/back-button.md)
    * [Код страницы с кнопками "Назад" и "Вперед"](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/back-button) 
* **Навигация.** Показывает меню (так называемое меню "гамбургер") для перехода между страницами в области задач. 
    * [Спецификация навигации](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/contextual-menu.md)
    * [Код навигации](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/navigation) 
* **Навигация с командами.** Показывает меню "гамбургер" с командными кнопками (или управляющими кнопками) в области задач. С помощью этого шаблона можно одновременно предоставлять параметры навигации и команды. 
    * [Спецификация навигации с командами](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/command-bar.md)
    * [Код навигации с командами](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/navigation-commands)
* **Сводка.** Показывает сводную панель навигации в области задач. С помощью сводной панели навигации пользователи смогут переключаться между различным контентом.
    * [Спецификация сводки](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/pivot.md)
    * [Код сводки](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/pivot)
* **Панель вкладок.** Показывает панель навигации, используя кнопки с расположенными по вертикали текстом и значками. Панель вкладок обеспечивает навигацию с помощью вкладок с короткими и понятными названиями.
    * [Спецификация панели вкладок](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/tab-bar.md)
    * [Код панели вкладок](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/tab-bar) 

<table>
<tr><th>Кнопка "Назад"</th><th>Навигация</th><th>Навигация с командами</th></tr>
<tr>
    <td>
        <A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/back-button">
        <img src="../images/back-button.png" alt="back button" style="width: 250px;"/></A>
    </td>
    <td>
        <A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/navigation">
        <img src="../images/navigation.png" alt="navigation" style="width: 250px;"/></A>
    </td>
    <td>
        <A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/navigation-commands">
        <img src="../images/navigation-commands.png" alt="navigation with commands" style="width: 250px;"/></A>
    </td>
</tr>
 </table>

<table>
<tr><th>Сводка</th><th>Панель вкладок</th></tr>
<tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/pivot">
<img src="../images/pivot.png" alt="pivot navigation" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/tab-bar">
<img src="../images/tab-bar.png" alt="tab bar" style="width: 250px;"/></A></td>
</tr>
 </table>

### <a name="notifications"></a>Уведомления

Надстройка может уведомлять пользователей о событиях, например ошибках, и о ходе выполнения множеством способов. Доступны следующие шаблоны уведомлений: 

* **Внедренное диалоговое окно.** Отображается в области задач и предоставляет сведения и (при необходимости) средства взаимодействия в виде кнопок и других элементов управления. Рекомендуется использовать диалоговое окно для подтверждения пользователем каких-либо действий. Шаблон "Внедренное диалоговое окно" позволяет ограничить пользовательский интерфейс областью задач.
    * [Спецификация внедренного диалогового окна](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/embedded-dialog.md)
    * [Код внедренного диалогового окна](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/embedded-dialog)
* **Встроенное сообщение.** Содержит информацию об ошибках, успешном выполнении действий или другие сведения. Может отображаться в указанном расположении в области задач. Например, если пользователь вводит в текстовом поле электронный адрес с неправильным форматом, то под полем отобразится соответствующее сообщение об ошибке. 
    * [Спецификация встроенного сообщения](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/notification_inlineMessage.pdf) (этот конструктивный шаблон заархивирован. Пока мы оцениваем его ценность, руководствуйтесь данным PDF-файлом.)
    * [Код встроенного сообщения](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/inline-message)
* **Баннер с сообщением.** Предоставляет сведения и (при необходимости) простые призывы к действиям в виде баннера, который можно свернуть в одну строку, развернуть на несколько строк или закрыть. Баннеры с сообщениями рекомендуется использовать для информирования об обновлениях служб или отображения полезных советов при запуске надстройки. 
    * [Спецификация баннера с сообщением](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/message_bar.pdf) (этот конструктивный шаблон заархивирован. Пока мы оцениваем его ценность, руководствуйтесь данным PDF-файлом.)
    * [Код баннера с сообщением](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/message-banner)
* **Индикатор выполнения.** Показывает ход выполнения длительных синхронных процессов, например задач по настройке, которые необходимо выполнить, прежде чем пользователь сможет перейти к дальнейшим действиям. Это отдельная промежуточная страница, на которой также используется фирменная символика надстройки. Используйте индикатор выполнения, если процесс может периодически отправлять сведения о том, сколько времени осталось до его завершения.
    * [Спецификация индикатора выполнения](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/progress-indicator.md)
    * [Код индикатора выполнения](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/progress-bar)
* **Индикатор работы.** Указывает, что выполняется длительный синхронный процесс, но не указывает, сколько времени осталось до его завершения. Это отдельная промежуточная страница, на которой также используется фирменная символика надстройки. Используйте индикатор работы, если надстройка не может достоверно сообщить, сколько времени необходимо для завершения процесса. 
    * [Спецификация вертушки](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/spinner.md)
    * [Код индикатора работы](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/spinner)
* **Всплывающее уведомление.** Содержит краткое сообщение, исчезающее через несколько секунд. Так как пользователь может и не увидеть такое сообщение, всплывающие уведомления используются для отображения несущественной информации. Это хороший способ уведомлять пользователей о событиях в удаленной системе, например о получении электронного письма.
    * [Спецификация всплывающего уведомления](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/toast.md)
    * [Код всплывающего уведомления](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/toast)

 <table>
 <tr><th>Внедренное диалоговое окно</th><th>Встроенное сообщение</th><th>Баннер с сообщением</th></tr>
 <tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/embedded-dialog"><img src="../images/embedded-dialogs.png" alt="embedded dialog" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/inline-message"><img src="../images/inline-messages.png" alt="inline message" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/message-banner"><img src="../images/message-banners.png" alt="message banner" style="width: 250px;"/></A></td></tr>
 </table>

 <table>
 <tr><th>Индикатор выполнения</th><th>Индикатор работы</th><th>Всплывающее уведомление</th></tr>
 <tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/progress-bar"><img src="../images/progress-bars.png" alt="progress bar" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/spinner"><img src="../images/logo-spinner.png" alt="spinner" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/toast"><img src="../images/toast-header.png" alt="toast" style="width: 250px;"/></A></td></tr>
 </table>
 


### <a name="general-components"></a>Общие компоненты

Ниже перечислены общие компоненты, которые можно использовать в надстройке для различных сценариев.  

#### <a name="client-dialog-boxes"></a>Диалоговые окна клиентов

Диалоговые окна клиентов обеспечивают еще один способ работы пользователей с надстройкой за пределами области задач. Доступны следующие шаблоны диалоговых окон:

* **Диалоговое окно с набором шрифтов.** Показывает диалоговое окно с текстовым содержимым. С помощью этого диалогового окна можно показывать пользователям подробные сведения. 
    * Узнайте больше о разработке [диалоговых окон в надстройках Office](dialog-boxes.md). Кроме того, следуйте нашим рекомендациям по [шрифтовому оформлению надстроек Office](add-in-design-language.md#typography).
    * [Код диалогового окна с набором шрифтов](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/typeramp)
* **Диалоговое окно с оповещением** содержит важные сведения для пользователей, например ошибки или уведомления.  
    * [Спецификация диалогового окна с оповещением](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/notification_alert.pdf) (этот конструктивный шаблон заархивирован. Пока мы оцениваем его ценность, руководствуйтесь данным PDF-файлом.)
    * [Код диалогового окна с оповещением](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/alert)
* **Диалоговое окно навигации**. Показывает диалоговое окно с элементами навигации. С его помощью пользователи могут переходить к различному содержимому. 
    * Узнайте больше о разработке [диалоговых окон в надстройках Office](dialog-boxes.md). Кроме того, узнайте, как использовать [сводные компоненты Office UI Fabric в надстройках Office](pivot.md).
    * [Код диалогового окна навигации](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/navigation)

<table>
 <tr><th>Диалоговое окно с набором шрифтов</th><th>Диалоговое окно с оповещением</th></tr>
<tr>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/typeramp"><img src="../images/typeramp-dialog.png" alt="typeramp dialog" width="400"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/alert"><img src="../images/alert-dialog.png" alt="alert dialog" width="400"/></A></td>
</tr></tr>
 </table>
 
 <table>
 <tr><th>Диалоговое окно навигации</th></tr>
<tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/navigation"><img src="../images/navigation-dialog.png" alt="navigation dialog" width="450"/></A></td></tr>
</tr>
 </table>


#### <a name="feedback-and-ratings"></a>Отзывы и оценки

Чтобы сделать надстройку более заметной и популярной, полезно предоставить пользователям возможность оценивать и рецензировать надстройку в AppSource. В этом шаблоне показано два способа предоставления отзывов и оценок из надстройки:

- Отзыв, инициированный пользователем, — пользователь отправляет отзыв с помощью меню навигации (например, ссылки **Отправить отзыв**) или значка в нижнем колонтитуле.
- Отзыв, инициированный системой, — после трех запусков надстройки появляется баннер с сообщением, в котором пользователю предлагается оставить отзыв.

Независимо от способа открывается диалоговое окно со страницей надстройки в AppSource.

* [Спецификация отзывов и оценок](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/notification_feedback.pdf) (этот конструктивный шаблон заархивирован. Пока мы оцениваем его ценность, руководствуйтесь данным PDF-файлом.)
* [Код отзывов и оценок](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/feedback/office-store)

> [!IMPORTANT]
> В настоящее время этот шаблон указывает на главную страницу AppSource. Обязательно замените эту ссылку на URL-адрес страницы вашей надстройки в AppSource.


 <table>
 <tr><th>Отзывы и оценки</th></tr>
<tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/feedback/office-store"><img src="../images/feedback-rating.png" alt="Feedback and Ratings" style="width: 350px;"/></A></td></tr>
</tr>
 </table>

#### <a name="settings-and-privacy"></a>Параметры и конфиденциальность

В надстройке может потребоваться страница параметров, где пользователи смогут настраивать параметры, контролирующие работу надстройки. Кроме того, вы можете сообщить пользователям о политиках конфиденциальности, соблюдаемых вашей надстройкой. 

* **Параметры.**. Показывает область задач с компонентами конфигурации, которая контролирует работу надстройки. На странице параметров пользователи могут выбрать нужные настройки.
    * [Спецификация параметров](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/settings.md)
    * [Код параметров](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/settings)
* **Политика конфиденциальности.** Показывает область задач с важными сведениями о политиках конфиденциальности. 
    * [Спецификация политики конфиденциальности](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/general_multiSection.pdf) (этот конструктивный шаблон заархивирован. Пока мы оцениваем его ценность, руководствуйтесь данным PDF-файлом.)
    * [Код политики конфиденциальности](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/settings)

<table>
 <tr><th>Параметры</th><th>Политика конфиденциальности</th></tr>
<tr>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/settings"><img src="../images/settings.png" alt="settings" style="width: 300px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/settings"><img src="../images/privacy-policy.png" alt="privacy" style="width: 264px;"/></A></td>
</tr></tr>
 </table>

## <a name="see-also"></a>См. также

* [Рекомендации по разработке надстроек Office](../concepts/add-in-development-best-practices.md)
* [Office UI Fabric](http://dev.office.com/fabric/)
