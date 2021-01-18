import { PreventIframe } from "express-msteams-host";

/**
 * Used as place holder for the decorators
 */
@PreventIframe("/agendaTab/index.html")
@PreventIframe("/agendaTab/config.html")
@PreventIframe("/agendaTab/remove.html")
export class AgendaTab {
}
