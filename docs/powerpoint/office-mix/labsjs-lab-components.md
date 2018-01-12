
# <a name="labsjs-lab-components"></a>LabsJS lab components

Labs.js предоставляет четыре типа компонентов, которые можно использовать для сборки лаборатории. Каждый тип компонента поддерживает определенный тип взаимодействия в лаборатории, например задачи с несколькими вариантами ответа, задачи с открытыми ответами или такие действия, как просмотр веб-страниц в iFrame урока в формате HTML.

## <a name="components"></a>Компоненты

Office Mix поддерживает следующие четыре типа компонентов лаборатории: 


-  **Компонент действия** (**IActivityComponent**). Указывает пользователю действие, которое необходимо выполнить, например прочитать отрывок текста, просмотреть видео или поработать с симуляторами. Дополнительные сведения см. в статье [Labs.Components.ActivityComponentInstance](../../../reference/office-mix/labs.components.activitycomponentinstance.md).
    
-  **Компонент выбора** (**IChoiceComponent**). Предоставляет пользователю список вариантов для выбора, из которых можно выбрать один или несколько ответов (или не отвечать). Этот компонент можно использовать в вопросах, предполагающих ответы "правда/неправда", выбор из нескольких вариантов и возможность выбрать несколько ответов, а также в опросах. Дополнительные сведения см. в статье [Labs.Components.ChoiceComponentInstance](../../../reference/office-mix/labs.components.choicecomponentinstance.md).
    
-  **Компонент ввода** (**IInputComponent**). Позволяет пользователю вводить данные в свободной форме. Этот тип компонента можно использовать, если нужно получить от пользователей ответы на поставленные вопросы или варианты решения математических задач, а также в тех случаях, когда пользователь в ответ на вопрос должен ввести какой-то текст. Дополнительные сведения см. в статье [Labs.Components.InputComponentInstance](../../../reference/office-mix/labs.components.inputcomponentinstance.md).
    
-  **Динамический компонент** (**IDynamicComponent**). Создает другие типы компонентов во время выполнения. Его можно использовать в вопросах, в которых следующий тип компонента отличается в зависимости от предыдущего ответа пользователя. Кроме того, этот тип компонента позволяет создавать банк вопросов для тестов и задач непосредственно во время выполнения. Дополнительные сведения см. в статье [Labs.Components.DynamicComponentInstance](../../../reference/office-mix/labs.components.dynamiccomponentinstance.md).
    

## <a name="additional-resources"></a>Дополнительные ресурсы



- [Надстройки Office Mix](../../powerpoint/office-mix/office-mix-add-ins.md)
    
- [Настройка и редактирование лабораторий LabsJS для Office Mix](../../powerpoint/office-mix/configuring-and-editing-labsjs-labs-for-office-mix.md)
    
- [Пошаговое руководство. Создание первой лаборатории для Office Mix](../../powerpoint/office-mix/creating-your-first-lab-for-office-mix.md#walkthrough-creating-your-first-lab-for-office-mix)
    
