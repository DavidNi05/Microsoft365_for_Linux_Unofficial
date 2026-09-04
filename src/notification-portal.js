const dbus =
    require("dbus-next");

const {
    Variant
} = dbus;


const PORTAL_DESTINATION =
    "org.freedesktop.portal.Desktop";

const PORTAL_PATH =
    "/org/freedesktop/portal/desktop";

const NOTIFICATION_INTERFACE =
    "org.freedesktop.portal.Notification";


let bus =
    null;

let notificationInterface =
    null;


/*
 * Get the XDG Notification Portal interface.
 */
async function getNotificationInterface() {

    if (
        notificationInterface
    ) {

        return notificationInterface;
    }


    if (
        !bus
    ) {

        bus =
            dbus.sessionBus();
    }


    const proxy =
        await bus.getProxyObject(
            PORTAL_DESTINATION,
            PORTAL_PATH
        );


    notificationInterface =
        proxy.getInterface(
            NOTIFICATION_INTERFACE
        );


    return notificationInterface;
}


/*
 * Creates a stable but unique notification ID.
 */
function createNotificationId(
    prefix = "m365"
) {

    return (
        prefix +
        "-" +
        Date.now() +
        "-" +
        Math.random()
            .toString(16)
            .slice(2)
    );
}


/*
 * Send a notification through XDG Desktop Portal.
 *
 * This works properly inside Flatpak and allows
 * Cinnamon/XApp to display the notification.
 */
async function showNotification({
    id = null,
    title,
    body,
    priority = "normal"
}) {

    try {

        const notifications =
            await getNotificationInterface();


        const notificationId =
            id ||
            createNotificationId();


        const payload = {

            title:
                new Variant(
                    "s",
                    String(
                        title ||
                        "Microsoft 365 for Linux"
                    )
                ),

            body:
                new Variant(
                    "s",
                    String(
                        body ||
                        ""
                    )
                ),

            priority:
                new Variant(
                    "s",
                    priority
                )
        };


        await notifications.AddNotification(
            notificationId,
            payload
        );


        console.log(
            "[Notification Portal] Notification sent:",
            notificationId
        );


        return {
            ok: true,
            id: notificationId
        };


    } catch (
        error
    ) {

        console.error(
            "[Notification Portal] Failed:",
            error
        );


        return {
            ok: false,
            error:
                error.message
        };
    }
}


/*
 * Remove a previously displayed notification.
 */
async function removeNotification(
    id
) {

    if (
        !id
    ) {
        return;
    }


    try {

        const notifications =
            await getNotificationInterface();


        await notifications.RemoveNotification(
            id
        );


    } catch (
        error
    ) {

        console.error(
            "[Notification Portal] Unable to remove notification:",
            error
        );
    }
}


module.exports = {

    showNotification,
    removeNotification
};
