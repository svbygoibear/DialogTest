export class DialogPopUp {
    private internalDialog: any;
    private internalOnOpen: () => void;
    private internalOnClose: () => void;

    get onOpen() {
        return this.internalOnOpen;
    }

    set onOpen(value: () => void) {
        this.internalOnOpen = value;
    }

    get onClose() {
        return this.internalOnClose;
    }

    set onClose(value: () => void) {
        this.internalOnClose = value;
    }

    public close(): void {
        if (
            typeof this.internalDialog !== "undefined" &&
            this.internalDialog !== null &&
            typeof this.internalDialog.close !== "undefined"
        ) {
            try {
                this.internalDialog.close();
                if (this.internalOnClose) {
                    this.internalOnClose();
                }
            } catch (e) {
                console.log(e);
            }
        }
    }

    public messageParent(message: string): void {
        Office.context.ui.messageParent(message);
    }

    public showDialog(
        options: any,
        messageHandler: (dialog: any, message: string) => void,
        eventHandler?: (dialog: any, message: string) => void
    ): void {
        console.log("show dialog");

        const fullUrl = options.route;

        console.log(fullUrl);
        Office.context.ui.displayDialogAsync(
            fullUrl,
            {
                height: options.heightPercent,
                width: options.widthPercent,
                displayInIframe: options.inFrame
            },
            (asyncResult: Office.AsyncResult<any>) => {
                const dialog = asyncResult.value;
                if (typeof dialog !== "undefined" && dialog !== null) {
                    this.internalDialog = dialog;
                    if (this.internalOnOpen) {
                        this.internalOnOpen();
                    }
                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg: any) => {
                        messageHandler(this, arg.message);
                    });
                    dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg: any) => {
                        if (eventHandler) {
                            eventHandler(this, arg);
                        }
                        if (this.internalOnClose && arg.error === 12006) {
                            this.internalOnClose();
                        }
                    });
                }
            }
        );
    }
}
