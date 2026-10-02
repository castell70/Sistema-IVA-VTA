import { html, nothing } from 'lit-html';
import { getAll, addEntity, updateEntity, deleteEntity } from '../data-store.js';
import { getFormData } from '../utils/form-utils.js';
import { forceRender } from '../router.js';

let entityType = 'clients'; // Default view type
let editingId = null; // State for tracking the currently edited entity ID
let updateNotification = null;

// Search/filter state: shared for both clients and suppliers
let searchQuery = '';

// NEW: Confirmation modal state for this view (avoid window.confirm)
let confirmMessage = null;
let confirmAction = null;
let confirmEntityId = null;

function showNotification(message, isSuccess = true) {
    updateNotification = { message, isSuccess };
    // We rely on the calling function (submit/save/cancel handler) to call forceRender 
    // to update the view immediately if structural changes occurred.
    setTimeout(() => {
        updateNotification = null;
        forceRender(); // Clear notification (asynchronous)
    }, 5000);
}

function showConfirmation(message, action, entityId) {
    confirmMessage = message;
    confirmAction = action;
    confirmEntityId = entityId;
    forceRender();
}

function closeConfirmation() {
    confirmMessage = null;
    confirmAction = null;
    confirmEntityId = null;
    forceRender();
}

function handleConfirmExecute() {
    if (confirmAction && confirmEntityId) confirmAction(confirmEntityId);
    closeConfirmation();
}

function ConfirmationModal() {
    if (!confirmMessage) return nothing;
    return html`
        <div class="modal-overlay" @click=${closeConfirmation}>
            <div class="modal-content" @click=${(e) => e.stopPropagation()}>
                <h3>Confirmación Requerida</h3>
                <p>${confirmMessage}</p>
                <div class="modal-actions">
                    <button class="cancel-btn" @click=${closeConfirmation} style="background-color: #6c757d;">Cancelar</button>
                    <button @click=${handleConfirmExecute} style="background-color: var(--error-color);">Confirmar</button>
                </div>
            </div>
        </div>
    `;
}


function handleFormSubmit(e) {
    e.preventDefault();
    const form = e.target;
    const data = getFormData(form);
    const isClient = entityType === 'clients';

    addEntity(entityType, data);
    
    // Clear form
    form.reset();
    
    // Set notification state
    showNotification(`${isClient ? 'Cliente' : 'Proveedor'} registrado exitosamente.`, true);

    // Force re-render of the current view to show the new entry and notification
    forceRender(); 
}

function handleEdit(e, entityId) {
    e.preventDefault();
    // Prevent editing if another row is already being edited
    if (editingId !== null) return; 
    editingId = entityId;
    // Force re-render to switch to edit mode
    forceRender(); 
}

function handleCancel(e) {
    e.preventDefault();
    editingId = null;
    // Force re-render to exit edit mode
    forceRender(); 
}

function handleSave(e, entityId) {
    e.preventDefault();
    const row = e.target.closest('tr');
    const isClient = entityType === 'clients';
    
    let notificationMessage = '';
    let notificationSuccess = false;

    try {
        const newData = {
            name: row.querySelector('[data-field="name"]').textContent.trim(),
            nrc: row.querySelector('[data-field="nrc"]').textContent.trim(),
            nit: row.querySelector('[data-field="nit"]').textContent.trim(),
            address: row.querySelector('[data-field="address"]').textContent.trim(),
            activity: row.querySelector('[data-field="activity"]').textContent.trim(),
            contact: row.querySelector('[data-field="contact"]').textContent.trim(),
            email: (row.querySelector('[data-field="email"]').textContent || '').trim(),
            // Normalize phone: keep only digits and limit to 10 characters
            phone: (row.querySelector('[data-field="phone"]').textContent || '').replace(/\D+/g, '').slice(0, 10),
        };
        
        if (!newData.nit) {
            throw new Error("El campo NIT es obligatorio y no puede ser vacío.");
        }
        
        // Validate phone if provided: allow empty or 7-10 digits
        if (newData.phone && (newData.phone.length < 7 || newData.phone.length > 10)) {
            throw new Error("Teléfono inválido: debe contener entre 7 y 10 dígitos.");
        }
        
        // Normalize NRC if empty
        newData.nrc = newData.nrc || null;


        updateEntity(entityType, entityId, newData);
        notificationMessage = `${isClient ? 'Cliente' : 'Proveedor'} actualizado exitosamente.`;
        notificationSuccess = true;

    } catch (error) {
        notificationMessage = `Error al actualizar: ${error.message}`;
        notificationSuccess = false;
    } finally {
        editingId = null; // Exit edit mode regardless of success/failure
        
        // Set notification state and schedule clear
        showNotification(notificationMessage, notificationSuccess);
        
        // Force re-render to update list/exit edit mode AND display notification
        forceRender(); 
    }
}

function handleDelete(entityId) {
    // Prevent deletion while editing another row
    if (editingId !== null) return showNotification("No puede eliminar mientras está editando otro registro.", false);

    const entities = getAll(entityType);
    const entity = entities.find(e => e.id === entityId);
    if (!entity) return showNotification("Registro no encontrado.", false);

    const name = entity.name || 'Registro';
    // Use in-UI confirmation modal instead of window.confirm
    showConfirmation(
        `ADVERTENCIA: ¿Desea eliminar definitivamente "${name}"? Esta acción no se puede deshacer.`,
        (id) => {
            try {
                const success = deleteEntity(entityType, id);
                if (success) {
                    editingId = null;
                    showNotification(`${name} eliminado correctamente.`, true);
                } else {
                    showNotification("No se pudo eliminar el registro. Intente nuevamente.", false);
                }
            } catch (err) {
                console.error("Error eliminado entidad:", err);
                showNotification("Ocurrió un error al eliminar. Revisa la consola.", false);
            } finally {
                forceRender();
            }
        },
        entityId
    );
}

// Search helper: filter by name, nrc or activity (case-insensitive substring)
function applySearchFilter(list) {
    if (!searchQuery || !searchQuery.trim()) return list;
    const q = searchQuery.trim().toLowerCase();
    return list.filter(item => {
        const name = (item.name || '').toLowerCase();
        const nrc = (item.nrc || '').toLowerCase();
        const activity = (item.activity || '').toLowerCase();
        return name.includes(q) || nrc.includes(q) || activity.includes(q);
    });
}

function handleSearchInput(e) {
    searchQuery = e.target.value || '';
    forceRender();
}

function EntityListTemplate(entities) {
    const filtered = applySearchFilter(entities);

    if (filtered.length === 0) {
        return html`<p>No hay ${entityType === 'clients' ? 'clientes' : 'proveedores'} registrados que coincidan.</p>`;
    }

    const isClient = entityType === 'clients';

    return html`
        <!-- Wrap the table in a fixed-height scroll area to present ~3 rows at a time -->
        <div style="max-height: 240px; overflow-y: auto; border: 1px solid var(--border-color); border-radius:6px; padding:6px; background:#fff;">
            <table class="data-table editable-table" style="margin:0;">
                <thead>
                    <tr>
                        <th>Nombre</th>
                        <th>NRC (${isClient ? 'Cliente' : 'Proveedor'})</th>
                        <th>NIT</th>
                        <th>Dirección</th>
                        <th>Actividad</th>
                        <th>Contacto</th>
                        <th>Teléfono</th>
                        <th>Correo</th>
                        <th>Acciones</th>
                    </tr>
                </thead>
                <tbody>
                    ${filtered.map(entity => {
                        const isEditing = entity.id === editingId;

                        return html`
                            <tr>
                                <td data-field="name" contenteditable=${isEditing}>${entity.name}</td>
                                <td data-field="nrc" contenteditable=${isEditing}>${entity.nrc || ''}</td>
                                <td data-field="nit" contenteditable=${isEditing}>${entity.nit}</td>
                                <td data-field="address" contenteditable=${isEditing}>${entity.address}</td>
                                <td data-field="activity" contenteditable=${isEditing}>${entity.activity}</td>
                                <td data-field="contact" contenteditable=${isEditing}>${entity.contact || ''}</td>
                                <td data-field="phone" contenteditable=${isEditing}>${entity.phone || ''}</td>
                                <td data-field="email" contenteditable=${isEditing}>${entity.email || ''}</td>
                                <td>
                                    ${isEditing 
                                        ? html`
                                            <button class="save-btn action-icon-btn" title="Guardar" @click=${(e) => handleSave(e, entity.id)}>💾</button>
                                            <button class="cancel-btn action-icon-btn" title="Cancelar" @click=${handleCancel} style="background-color: var(--error-color);">✖</button>
                                        `
                                        : html`
                                            <button class="edit-btn action-icon-btn" title="Editar" @click=${(e) => handleEdit(e, entity.id)} ?disabled=${editingId !== null}>✎</button>
                                            <button class="delete-btn action-icon-btn" title="Eliminar" @click=${() => handleDelete(entity.id)} ?disabled=${editingId !== null} style="color: var(--error-color);">🗑️</button>
                                        `
                                    }
                                </td>
                            </tr>
                        `;
                    })}
                </tbody>
            </table>
        </div>
    `;
}

function EntityFormTemplate() {
    const isClient = entityType === 'clients';
    const title = isClient ? 'Registrar Nuevo Cliente' : 'Registrar Nuevo Proveedor';
    const nrcLabel = isClient ? 'NRC (Registro de Contribuyente - Obligatorio para CCF)' : 'NRC Proveedor (Obligatorio para CCF)';

    return html`
        <div class="card">
            <h3>${title}</h3>
            <form @submit=${handleFormSubmit}>

                <!-- Row: Name full width -->
                <div class="form-group">
                    <label for="name">Nombre / Razón Social</label>
                    <input type="text" id="name" name="name" required>
                </div>

                <!-- Row: NRC | NIT -->
                <div class="dashboard-grid" style="grid-template-columns: 1fr 1fr; gap: 15px;">
                    <div class="form-group">
                        <label for="nrc">${nrcLabel}</label>
                        <input type="text" id="nrc" name="nrc">
                    </div>
                    <div class="form-group">
                        <label for="nit">NIT (Número de Identificación Tributaria)</label>
                        <input type="text" id="nit" name="nit" required>
                    </div>
                </div>

                <!-- Row: Address full width -->
                <div class="form-group">
                    <label for="address">Dirección</label>
                    <textarea id="address" name="address" rows="2"></textarea>
                </div>

                <!-- Row: Activity | Phone | Contact -->
                <div class="dashboard-grid" style="grid-template-columns: 1fr 1fr 1fr; gap: 15px; align-items: end;">
                    <div class="form-group" style="margin-bottom:0;">
                        <label for="activity">Giro o Actividad Económica</label>
                        <input type="text" id="activity" name="activity">
                    </div>
                    <div class="form-group" style="margin-bottom:0;">
                        <label for="phone">Teléfono</label>
                        <input type="text" id="phone" name="phone" maxlength="10" pattern="\\d{7,10}" title="Ingrese entre 7 y 10 dígitos" placeholder="1234567890">
                    </div>
                    <div class="form-group" style="margin-bottom:0;">
                        <label for="contact">Contacto</label>
                        <input type="text" id="contact" name="contact">
                    </div>
                </div>

                <!-- Row: Email stretched and visually taller to match other controls -->
                <div class="dashboard-grid" style="grid-template-columns: 1.4fr; gap: 15px; margin-top:10px;">
                    <div class="form-group" style="margin-bottom:0;">
                        <label for="email">Correo Electrónico</label>
                        <!-- Increased width and maintained internal height so it visually matches other inputs -->
                        <input 
                            type="email" 
                            id="email" 
                            name="email" 
                            placeholder="correo@ejemplo.com" 
                            style="height:40px; padding-top:8px; padding-bottom:8px; width:100%; max-width:720px; display:block;"
                        >
                    </div>
                </div>

                <!-- Save button placed in its own row and lowered slightly for breathing room -->
                <div style="margin-top:18px; display:flex; justify-content:flex-start;">
                    <button type="submit">Guardar ${isClient ? 'Cliente' : 'Proveedor'}</button>
                </div>
            </form>
        </div>
    `;
}


export function EntitiesView(type) {
    entityType = type;
    const entities = getAll(entityType);
    const title = type === 'clients' ? 'Control de Clientes' : 'Control de Proveedores';

    return html`
        ${ConfirmationModal()}
        <h2>${title}</h2>

        ${updateNotification ? html`
            <div class="message ${updateNotification.isSuccess ? 'success' : 'error'}">
                ${updateNotification.message}
            </div>
        ` : nothing}
        
        ${EntityFormTemplate()}

        <div class="card" style="margin-top: 20px;">
            <h3>Listado de ${title}</h3>

            <!-- Search box: filters by name, NRC, activity -->
            <div style="margin-bottom: 12px; display:flex; gap:8px; align-items:center; flex-wrap:wrap;">
                <input 
                    type="search" 
                    placeholder="Buscar por Nombre, NRC o Actividad..." 
                    @input=${handleSearchInput} 
                    value="${searchQuery}"
                    style="flex:1; padding:8px; border:1px solid var(--border-color); border-radius:6px;"
                    aria-label="Buscar en listado"
                />
                <button type="button" @click=${() => { searchQuery=''; forceRender(); }} style="background-color:var(--accent-color);">Borrar</button>
            </div>

            ${EntityListTemplate(entities)}
        </div>
    `;
}