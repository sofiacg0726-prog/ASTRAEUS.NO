
console.log("ASTRAEUS: menú conectado");

document.addEventListener("DOMContentLoaded", function () {

    const boton = document.getElementById("menuBoton");
    const lateral = document.getElementById("menuLateral");
    const overlay = document.getElementById("menuOverlay");
    const cerrar = document.getElementById("cerrarMenu");

    if (!boton || !lateral || !overlay || !cerrar) {
        console.error("ASTRAEUS: Faltan elementos del menú");
        return;
    }

    function abrirMenu() {
        lateral.classList.add("menu-abierto");
        overlay.classList.add("overlay-visible");
        document.body.classList.add("menu-activo");

        boton.setAttribute("aria-expanded", "true");
        lateral.setAttribute("aria-hidden", "false");
    }

    function cerrarMenu() {
        lateral.classList.remove("menu-abierto");
        overlay.classList.remove("overlay-visible");
        document.body.classList.remove("menu-activo");

        boton.setAttribute("aria-expanded", "false");
        lateral.setAttribute("aria-hidden", "true");
    }

    boton.setAttribute("aria-expanded", "false");
    lateral.setAttribute("aria-hidden", "true");

    boton.addEventListener("click", abrirMenu);
    cerrar.addEventListener("click", cerrarMenu);
    overlay.addEventListener("click", cerrarMenu);

    document.addEventListener("keydown", function (evento) {
        if (evento.key === "Escape") {
            cerrarMenu();
        }
    });

    console.log("ASTRAEUS: menú funcionando correctamente");

});
