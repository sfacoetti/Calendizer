window.calendizer = (function () {
    const START = 0; // la colonna copre 0:00-24:00; gli slot hanno altezze diverse (ore fuori orario ridotte), quindi si leggono i rettangoli reali
    let dotnet = null, drag = null, ghost = null;

    function cellAt(x, y) {
        return document.elementsFromPoint(x, y).find(e => e.dataset && e.dataset.day);
    }
    // indice dello slot sotto y: con dy (grab dell'attività trascinata) si arrotonda, senza si prende lo slot puntato
    function slotRects(col) { return Array.from(col.children).filter(c => c.classList.contains('slot')).map(c => c.getBoundingClientRect()); }
    function slotIndex(col, y, dy) {
        const rs = slotRects(col);
        if (dy === undefined) { const i = rs.findIndex(r => y < r.bottom); return i < 0 ? rs.length - 1 : i; }
        const t = y - dy; // slot il cui bordo alto è il più vicino alla cima dell'attività trascinata
        let best = 0;
        rs.forEach((r, i) => { if (Math.abs(r.top - t) < Math.abs(rs[best].top - t)) best = i; });
        return best;
    }

    // "yyyy-MM-dd|minuti" dello slot sotto il punto (coordinate CSS del viewport), "|-1" = vista mese
    function slotAt(x, y, dy) {
        const el = cellAt(x, y);
        if (!el) return null;
        if (el.dataset.month) return el.dataset.day + "|-1";
        return el.dataset.day + "|" + (START + slotIndex(el, y, dy) * 30);
    }

    function hideGhost() { if (ghost) { ghost.remove(); ghost = null; } }

    // Anteprima della destinazione: rettangolo che occupa lo slot e si divide lo spazio con le attività già presenti.
    // Sta nel body (fixed) e non nella griglia, per non interferire con il rendering di Blazor.
    function showGhost(x, y) {
        const el = cellAt(x, y);
        if (!el) return hideGhost();
        if (!ghost) { ghost = document.createElement('div'); document.body.appendChild(ghost); }
        const r = el.getBoundingClientRect();
        const set = (l, t, w, h) => Object.assign(ghost.style, { left: l + 'px', top: t + 'px', width: w + 'px', height: h + 'px' });
        if (el.dataset.month) { ghost.className = 'ghost cell'; return set(r.left, r.top, r.width, r.height); }
        ghost.className = 'ghost';
        const rs = slotRects(el), n = (drag && drag.slots) || 1;
        const i = Math.min(slotIndex(el, y, drag && drag.dy), rs.length - n);
        const top = rs[i].top, h = rs[i + n - 1].bottom - top - 2;
        let k = 0; // attività già presenti nello stesso intervallo
        el.querySelectorAll('.act').forEach(a => {
            if (drag && a.dataset.id === drag.id) return;
            const b = a.getBoundingClientRect();
            if (b.top < top + h && b.bottom > top) k++;
        });
        set(r.left + r.width * k / (k + 1), top, r.width / (k + 1), h);
    }

    document.addEventListener('dragstart', e => {
        const a = e.target.closest && e.target.closest('[data-id]');
        if (!a) return;
        const isAct = a.classList.contains('act');
        drag = { id: a.dataset.id, dy: isAct ? e.clientY - a.getBoundingClientRect().top : undefined };
        if (isAct) drag.slots = Math.max(1, Math.round(+a.dataset.min / 30));
        e.dataTransfer.setData('text/plain', 'act');
        e.dataTransfer.effectAllowed = 'move';
    });
    // dragover sempre accettato: così anche il drop esterno (Outlook) arriva alla pagina invece di aprire un file
    document.addEventListener('dragover', e => {
        e.preventDefault();
        e.dataTransfer.dropEffect = drag ? 'move' : 'copy';
        showGhost(e.clientX, e.clientY);
    });
    document.addEventListener('dragleave', e => { if (!e.relatedTarget) hideGhost(); });
    document.addEventListener('drop', e => {
        e.preventDefault();
        hideGhost();
        if (!drag) { // drop esterno (mail da Outlook): lo slot lo calcoliamo qui, la mail la legge l'host
            const x = slotAt(e.clientX, e.clientY);
            if (x && dotnet) dotnet.invokeMethodAsync('External', x);
            return;
        }
        const s = slotAt(e.clientX, e.clientY, drag.dy), id = drag.id;
        drag = null;
        if (s && dotnet) dotnet.invokeMethodAsync('Move', id, s);
    });
    // Canc fuori dai campi di testo: l'app decide cosa eliminare
    document.addEventListener('keydown', e => {
        if (e.key !== 'Delete' || /^(INPUT|TEXTAREA|SELECT)$/.test(e.target.tagName) || !dotnet) return;
        e.preventDefault();
        dotnet.invokeMethodAsync('DeleteKey');
    });
    // Invio fuori dai campi: apre il popup sulle caselle selezionate
    document.addEventListener('keydown', e => {
        if (e.key !== 'Enter' || /^(INPUT|TEXTAREA|SELECT|BUTTON)$/.test(e.target.tagName) || !dotnet) return;
        dotnet.invokeMethodAsync('EnterKey');
    });
    // Esc: chiude il popup aperto (anche con il focus in un campo)
    document.addEventListener('keydown', e => {
        if (e.key === 'Escape' && dotnet) dotnet.invokeMethodAsync('EscKey');
    });
    document.addEventListener('dragend', () => { drag = null; hideGhost(); });

    document.addEventListener('pointerdown', e => {
        const h = e.target.closest && e.target.closest('.rz');
        if (!h) return;
        e.preventDefault(); e.stopPropagation();
        const a = h.closest('[data-id]');
        a.draggable = false;
        h.setPointerCapture(e.pointerId);
        const rs = slotRects(a.parentElement), s0 = Math.round(+a.dataset.start / 30), top = rs[s0].top, h0 = a.style.height;
        const slots = ev => { // numero di slot fino al bordo basso più vicino al puntatore
            let k = s0 + 1;
            for (let j = s0 + 1; j <= rs.length; j++) if (Math.abs(rs[j - 1].bottom - ev.clientY) < Math.abs(rs[k - 1].bottom - ev.clientY)) k = j;
            return k - s0;
        };
        const move = ev => { a.style.height = (rs[s0 + slots(ev) - 1].bottom - top - 2) + 'px'; };
        const up = ev => {
            h.removeEventListener('pointermove', move);
            h.removeEventListener('pointerup', up);
            a.draggable = true; a.style.height = h0; // lo stile torna a quello di Blazor (%), che si aggiorna da solo
            if (dotnet) dotnet.invokeMethodAsync('Resize', a.dataset.id, slots(ev) * 30);
        };
        h.addEventListener('pointermove', move);
        h.addEventListener('pointerup', up);
    });

    // Tema: "light" / "dark" come attributo sulla radice, "auto" lo toglie (il CSS segue Windows)
    function setTheme(t) {
        if (t === "light" || t === "dark") document.documentElement.dataset.theme = t;
        else delete document.documentElement.dataset.theme;
    }
    // Porta la vista alle 8:00 (la griglia mostra 8:00-18:30, il resto sta sopra/sotto)
    function scrollToDay() {
        const g = document.querySelector(".grid"), b = g && g.querySelector(".body");
        if (b) g.scrollTop = b.offsetHeight * 16 / 48;
    }
    return { init: ref => { dotnet = ref; }, slotAt, setTheme, scrollToDay };
})();
