const dbus = require("dbus-next");

const PORTAL_BUS_NAME =
    "org.freedesktop.portal.Desktop";

const PORTAL_OBJECT_PATH =
    "/org/freedesktop/portal/desktop";

const BACKGROUND_INTERFACE =
    "org.freedesktop.portal.Background";

const REQUEST_INTERFACE =
    "org.freedesktop.portal.Request";


function makeToken() {
    return (
        "m365_" +
        Date.now().toString(36) +
        "_" +
        Math.random()
            .toString(36)
            .slice(2, 10)
    ).replace(
        /[^A-Za-z0-9_]/g,
        "_"
    );
}


function unwrapVariant(value) {
    if (
        value &&
        typeof value === "object" &&
        Object.prototype.hasOwnProperty.call(
            value,
            "value"
        )
    ) {
        return value.value;
    }

    return value;
}


async function getBackgroundPortal() {
    const bus =
        dbus.sessionBus();

    const portalObject =
        await bus.getProxyObject(
            PORTAL_BUS_NAME,
            PORTAL_OBJECT_PATH
        );

    const background =
        portalObject.getInterface(
            BACKGROUND_INTERFACE
        );

    return {
        bus,
        background
    };
}


async function requestBackgroundAccess() {
    let bus = null;

    try {
        const portal =
            await getBackgroundPortal();

        bus =
            portal.bus;

        const background =
            portal.background;

        const token =
            makeToken();

        const options = {
            handle_token:
                new dbus.Variant(
                    "s",
                    token
                ),

            reason:
                new dbus.Variant(
                    "s",
                    "Keep Microsoft OneDrive files synchronized in the background."
                ),

            autostart:
                new dbus.Variant(
                    "b",
                    true
                ),

            commandline:
                new dbus.Variant(
                    "as",
                    [
                        "m365-linux",
                        "--background"
                    ]
                )
        };

        /*
         * An empty parent_window is valid.
         * This keeps the implementation desktop-neutral
         * and works on X11 and Wayland.
         */
        const requestPath =
            await background.RequestBackground(
                "",
                options
            );

        console.log(
            "[Background Portal] Request:",
            requestPath
        );

        const requestObject =
            await bus.getProxyObject(
                PORTAL_BUS_NAME,
                requestPath
            );

        const request =
            requestObject.getInterface(
                REQUEST_INTERFACE
            );

        return await new Promise(
            resolve => {

                let finished = false;

                const finish =
                    result => {

                        if (finished) {
                            return;
                        }

                        finished = true;

                        clearTimeout(
                            timeout
                        );

                        resolve(result);
                    };


                const timeout =
                    setTimeout(
                        () => {

                            console.error(
                                "[Background Portal] Request timed out"
                            );

                            finish({
                                granted: false,
                                background: false,
                                autostart: false,
                                response: -1,
                                timeout: true
                            });

                        },
                        120000
                    );


                request.on(
                    "Response",
                    (
                        response,
                        results
                    ) => {

                        console.log(
                            "[Background Portal] Response:",
                            response,
                            results
                        );

                        const backgroundAllowed =
                            Boolean(
                                unwrapVariant(
                                    results?.background
                                )
                            );

                        const autostartAllowed =
                            Boolean(
                                unwrapVariant(
                                    results?.autostart
                                )
                            );

                        finish({
                            /*
                             * 0 = accepted
                             * 1 = cancelled
                             * 2 = other denial/error
                             */
                            granted:
                                response === 0,

                            background:
                                backgroundAllowed,

                            autostart:
                                autostartAllowed,

                            response,

                            timeout: false
                        });
                    }
                );

            }
        );

    } catch (error) {

        console.error(
            "[Background Portal] Error:",
            error
        );

        return {
            granted: false,
            background: false,
            autostart: false,
            response: -1,
            timeout: false,
            error:
                error.message
        };

    } finally {

        /*
         * Do not disconnect the session bus here.
         * dbus-next can share it with other portal
         * operations while the app remains open.
         */

    }
}


async function setBackgroundStatus(
    message
) {
    try {

        const {
            background
        } =
            await getBackgroundPortal();

        await background.SetStatus({
            message:
                new dbus.Variant(
                    "s",
                    String(message)
                        .substring(
                            0,
                            96
                        )
                )
        });

        return true;

    } catch (error) {

        console.error(
            "[Background Portal] SetStatus failed:",
            error
        );

        return false;
    }
}


module.exports = {
    requestBackgroundAccess,
    setBackgroundStatus
};
