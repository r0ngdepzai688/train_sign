# IQC Nexus — 2026 Login UI/UX Refinement Design (Phương án A)

**Date:** 2026-10-27  
**Author:** Nguyễn Hải Đường (`hai.duong`)  
**Status:** Approved

---

## 1. Visual Flaws Addressed

1. **Footer Collision & 1400px Separator Line:**
   - Removed the center `<g class="circuit-socket">` dashed rectangle and crosshairs that collided with the footer copyright text.
   - Removed the full-width `border-top` on `.nexus-footer-brand` and constrained `.nexus-footer-brand` to match the card rhythm (`margin-top: 24px`).
2. **Sub-Pixel Semiconductor Circuit Layer & GPU Optical Data Pulses:**
   - Refined PCB trace stroke widths (`1px–1.25px`), softened opacity, added a radial vignette mask (`mask-image: radial-gradient(...)`) so the center login zone remains optically calm, and added subtle GPU-driven optical data pulses (`stroke-dashoffset` animation).
3. **Typography & Brand Hierarchy:**
   - Removed the redundant `IQC NEXUS • QUANTUM CORE` pill directly below the 3D `IQC NEXUS` metallic logo.
   - Switched `.nexus-copyright` from `JetBrains Mono` to `Plus Jakarta Sans` (`--font-sans`) for crisp Vietnamese diacritics on `"Bản quyền thuộc về Nguyễn Hải Đường (hai.duong)"`.
4. **Unified Input Shells & 12px Tactile Enterprise CTA:**
   - Replaced the 3-piece Bootstrap `.input-group` with `.nexus-input-shell` (single unified `12px` radius enclosure with internal icon and eye-toggle button, smooth `:focus-within` halo).
   - Aligned the primary Login/Register CTA button radius to `12px` (`--radius-btn`) with tactile top-edge specular highlight and arrow micro-motion.
   - Added a floating Light/Dark theme toggle button on `#view-login` and `#view-reg`.
