import { PlElement, css, html } from "polylib";

class PlDocumentExportTemplate extends PlElement {
    static properties = {
        /**
         * Текст
         */
        label: {
            type: String
        },
        /**
         * Текст
         */
        text: {
            type: String,
            reflectToAttribute: true
        },

        /**
         * Тип шаблона
         * * **title** -  Заголовок;
         * * **new-line** - Пустая строчка;
         * * **filter** - Фильтры (пустой аргумент text не отобразит строчку фильтра).
         */
        type: {
            type: String
        },

        /**
         * Позиционирование текста (только для type="filter"):
         * * **left** - слева;
         * * **center** - по центру;
         * * **right** - справа.
         */
        align: {
            type: String,
            value: 'left'
        },

        /**
         * Размер текста (только для type="title"):
         * * **h1** - 14px PDF, 12px XLS
         * * **h2** - 12px PDF, 10px XLS
         * * **h3** - 10px PDF, 9px XLS
         * * **h4** - 8px PDF, 8px XLS
         */
        size: {
            type: String
        }
    };

    static css = css`
        :host {
            display: none;
        }
    `;

}

customElements.define('pl-document-export-template', PlDocumentExportTemplate);
