---
title: Тестирование удобства использования надстроек Office
description: Узнайте, как протестировать структуру надстройки с реальными пользователями.
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 49a2af983615779160886961e8269e4588d0fc9e
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810283"
---
# <a name="usability-testing-for-office-add-ins"></a>Тестирование удобства использования надстроек Office

A great add-in design takes user behaviors into account. Because your own preconceptions influence your design decisions, it’s important to test designs with real users to make sure that your add-ins work well for your customers.

Тесты удобства использования можно выполнять разными способами. Для многих разработчиков надстроек удаленные и немодерированные исследования удобства использования являются наиболее экономичными и временными. Несколько популярных служб тестирования упрощают эту задачу; Ниже приведены некоторые примеры.

- [UserTesting.com](https://www.UserTesting.com)
- [Optimalworkshop.com](https://www.Optimalworkshop.com)
- [Userzoom.com](https://www.Userzoom.com)

Эти службы тестирования помогают упростить создание плана тестирования, а также избавляют от необходимости искать участников и наблюдать за тестированием.

You need only five participants to uncover most usability issues in your design. Incorporate small tests regularly throughout your development cycle to ensure that your product is user-centered.

> [!NOTE]
> We recommend that you test the usability of your add-in across multiple platforms. To [publish your add-in to AppSource](/office/dev/store/submit-to-appsource-via-partner-center), it must work on all [platforms that support the methods that you define](/javascript/api/requirement-sets).

## <a name="1-sign-up-for-a-testing-service"></a>1. Регистрация в службе тестирования

Дополнительные сведения см. в статье [Выбор веб-инструмента для удаленного немодерируемого тестирования](https://www.nngroup.com/articles/unmoderated-user-testing-tools/).

## <a name="2-develop-your-research-questions"></a>2. Определите предметы исследования

Research questions define the objectives of your research and guide your test plan. Your questions will help you identify participants to recruit and the tasks they will perform. Make your research questions as specific as you can. You can also seek to answer broader questions.

Ниже приведены некоторые примеры исследовательских вопросов.

**Конкретные**

- Замечают ли пользователи ссылку "Бесплатная пробная версия" на целевой странице?
- Когда пользователи вставляют содержимое из надстройки в документ, знают ли они, в каком месте документа оно будет вставлено?

**Общие**

- С чем у пользователя возникает больше всего сложностей при работе с надстройкой?
- Понимают ли пользователи значения значков на панели команд, прежде чем нажимать их?
- Легко ли пользователям найти меню настроек?

Очень важно собрать данные обо всем процессе работы пользователя — от обнаружения надстройки до ее установки и использования. Рассмотрите исследовательские вопросы, касающиеся следующих аспектов пользовательского интерфейса надстройки.

- поиск надстройки в AppSource;
- установка надстройки;
- первый запуск;
- команды ленты;
- интерфейс надстройки;
- взаимодействие надстройки с пространством документа в приложении Office;
- возможности пользователя по управлению вставкой содержимого.

Дополнительные сведения см. в статье [Gathering factual responses vs. subjective data](https://help.usertesting.com/hc/articles/115003378572-Writing-effective-questions) (Сбор фактических ответов и субъективных данных).

## <a name="3-identify-participants-to-target"></a>3. Определите участников тестирования

Remote testing services can give you control over many characteristics of your test participants. Think carefully about what kinds of users you want to target. In your early stages of data collection, it might be better to recruit a wide variety of participants to identify more obvious usability issues. Later, you might choose to target groups like advanced Office users, particular occupations, or specific age ranges.

## <a name="4-create-the-participant-screener"></a>4. Создайте анкету для отбора участников

The screener is the set of questions and requirements you will present to prospective test participants to screen them for your test. Keep in mind that participants for services like UserTesting.com have a financial interest in qualifying for your test. It's a good idea to include trick questions in your screener if you want to  exclude certain users from the test. 

Например, если вы хотите найти участников, знакомых с сайтом GitHub, добавьте в список вариантов ответа вымышленные названия, чтобы отфильтровать пользователей, предоставляющих ложные сведения.

**С каким из перечисленных ниже репозиториев кода вы знакомы?**  
 a. SourceShelf  [*Reject*]  
 b. CodeContainer  [*Reject*]  
 c. GitHub  [*Must select*]  
 d. BitBucket  [*May select*]  
 e. CloudForge  [*May select*]  

Если вы планируете тестировать готовую сборку надстройки, представленные ниже вопросы помогут вам отобрать пользователей, которые смогут это сделать.

**Для этого теста вам потребуется последняя версия Microsoft PowerPoint. У вас есть последняя версия PowerPoint?**  
 a. Yes [*Must select*]  
 b. No [*Reject*]  
 c. I don’t know [*Reject*]  

**Для этого теста вам потребуется установить бесплатную надстройку для PowerPoint и создать бесплатную учетную запись для ее использования. Вы готовы установить надстройку и создать бесплатную учетную запись?**  
 a. Yes [*Must select*]  
 b. No [*Reject*]  

Дополнительные сведения см. в статье [Рекомендации по составлению вопросов для отбора](https://help.usertesting.com/hc/articles/115003370731-Screener-question-best-practices).

## <a name="5-create-tasks-and-questions-for-participants"></a>5. Составьте список задач и вопросов для участников

Try to prioritize what you want tested so that you can limit the number of tasks and questions for the participant. Some services pay participants only for a set amount of time, so you want to make sure not to go over.

Try to observe participant behaviors instead of asking about them, whenever possible. If you need to ask about behaviors, ask about what participants have done in the past, rather than what they would expect to do in a situation. This tends to give more reliable results.

The main challenge in unmoderated testing is making sure your participants understand your tasks and scenarios. Your directions should be *clear and concise*. Inevitably, if there is potential for confusion, someone will be confused.

Don't assume that your user will be on the screen they’re supposed to be on at any given point during the test. Consider telling them what screen they need to be on to start the next task.

Дополнительные сведения см. в статье [Составление хороших задач](https://help.usertesting.com/hc/articles/115003371651-Writing-great-tasks).

## <a name="6-create-a-prototype-to-match-the-tasks-and-questions"></a>6. Создайте прототип, соответствующий задачам и вопросам

Вы можете тестировать либо готовую версию надстройки, либо ее прототип. Помните, что если вы хотите протестировать готовую версию надстройки, вам потребуется отобрать участников, у которых есть последняя версия Office и которые готовы установить надстройку и зарегистрировать учетную запись (если у вас нет для них готовых учетных данных для входа). Затем необходимо убедиться, что они успешно установили надстройку.

On average, it takes about 5 minutes to walk users through how to install an add-in. The following is an example of clear, concise installation steps. Adjust the steps based on the specifics of your test.

**Установите надстройку (вставьте здесь имя надстройки) для PowerPoint, следуя приведенным ниже инструкциям.**

1. Откройте Microsoft PowerPoint.
1. Выберите элемент **Пустая презентация**.
1. Перейдите к **разделу Вставка** > **надстроек**.
1. Во всплывающем окне выберите **Магазин**.
1. Введите <название надстройки> в поле поиска.
1. Выберите <название надстройки>.
1. Изучите страницу Магазина, чтобы ознакомиться с надстройкой.
1. Нажмите **Добавить**, чтобы установить надстройку.

You can test a prototype at any level of interaction and visual fidelity. For more complex linking and interactivity, consider a prototyping tool like [InVision](https://www.invisionapp.com). If you just want to test static screens, you can host images online and send participants the corresponding URL, or give them a link to an online PowerPoint presentation. 

## <a name="7-run-a-pilot-test"></a>7. Проведите пилотный тест

It can be tricky to get the prototype and your task/question list right. Users might be confused by tasks, or might get lost in your prototype. You should run a pilot test with 1-3 users to work out the inevitable issues with the test format. This will help to ensure that your questions are clear, that the prototype is set up correctly, and that you’re capturing the type of data you’re looking for.

## <a name="8-run-the-test"></a>8. Проведите тестирование

After you order your test, you will get email notifications when participants complete it. Unless you’ve targeted a specific group of participants, the tests are usually completed within a few hours.

## <a name="9-analyze-results"></a>9. Проанализируйте результаты

This is the part where you try to make sense of the data you’ve collected. While watching the test videos, record notes about problems and successes the user has. Avoid trying to interpret the meaning of the data until you have viewed all the results.

A single participant having a usability issue is not enough to warrant making a change to the design. Two or more participants encountering the same issue suggests that other users in the general population will also encounter that issue.

In general, be careful about how you use your data to draw conclusions. Don’t fall into the trap of trying to make the data fit a certain narrative; be honest about what the data actually proves, disproves, or simply fails to provide any insight about. Keep an open mind; user behavior frequently defies designer’s expectations.

## <a name="see-also"></a>См. также

- [Как проводить тестирование удобства использования](https://whatpixel.com/howto-conduct-usability-testing/)  
- [Рекомендации по пользовательскому тестированию](https://help.usertesting.com/hc/articles/115003370231-Best-practices-for-UserTesting)  
- [Предотвращение предвзятости](https://downloads.usertesting.com/white_papers/TipSheet_MinimizingBias.pdf)  
