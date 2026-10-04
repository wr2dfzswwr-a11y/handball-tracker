import React, { useEffect, useRef } from "react";

// Keep the game mounted while help is open: opening help must not reset its clock.
export default function GuideDialog({ onClose }) {
  const ref = useRef(null);
  useEffect(() => {
    const dialog = ref.current;
    const previousFocus = document.activeElement;
    const previousOverflow = document.body.style.overflow;
    dialog.showModal();
    document.body.style.overflow = "hidden";
    return () => {
      document.body.style.overflow = previousOverflow;
      if (dialog.open) dialog.close();
      previousFocus?.focus?.();
    };
  }, []);
  return (
    <dialog ref={ref} className="guide-dialog" aria-labelledby="guide-title"
      onCancel={onClose} onClose={onClose}>
      <div className="guide-dialog-bar">
        <strong id="guide-title">Anleitung</strong>
        <button type="button" onClick={onClose} aria-label="Anleitung schließen">✕ Schließen</button>
      </div>
      <iframe title="Handballtracker – vollständige Anleitung" src="./guide.html" />
    </dialog>
  );
}
