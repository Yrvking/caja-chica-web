# Guía: Actualizar Reglas de Firebase para Caja Chica

## Firestore Rules

1. Ir a [Firebase Console](https://console.firebase.google.com/) → proyecto **planilla-webg**
2. En el menú izquierdo: **Firestore Database** → pestaña **Reglas**
3. Reemplaza **todo el contenido** con el bloque de abajo
4. Clic en **Publicar**

```
rules_version = '2';
service cloud.firestore {
  match /databases/{database}/documents {

    function isSignedIn() {
      return request.auth != null;
    }

    function userData() {
      return isSignedIn()
        ? get(/databases/$(database)/documents/users/$(request.auth.uid)).data
        : null;
    }

    function isAdmin() {
      return isSignedIn() && userData().rol == 'ADMIN';
    }

    function isSelf(userId) {
      return isSignedIn() && request.auth.uid == userId;
    }

    function isApproverForArea(areaId) {
      return isSignedIn()
        && areaId != null
        && exists(/databases/$(database)/documents/areas/$(areaId))
        && get(/databases/$(database)/documents/areas/$(areaId)).data.approverUid == request.auth.uid;
    }

    function onlyVbFieldsChanged() {
      return request.resource.data.diff(resource.data).affectedKeys()
        .hasOnly(['vbOk','vbByUid','vbByName','vbAt']);
    }

    function comprobantesWriteOk() {
      return !request.resource.data.diff(resource.data).affectedKeys().hasAny(['comprobantes'])
        || request.resource.data.pagado == true
        || (request.resource.data.comprobantes is list && request.resource.data.comprobantes.size() == 0);
    }

    // ═══════════════════════════════════
    //  COMPARTIDAS (Planilla + Caja Chica + Movilidad)
    // ═══════════════════════════════════

    match /users/{userId} {
      allow read: if isSelf(userId) || isAdmin();
      allow create: if isAdmin();
      allow update: if isAdmin()
        || (isSelf(userId) && !request.resource.data.diff(resource.data).affectedKeys().hasAny(['rol','areaId']));
      allow delete: if isAdmin();
    }

    match /empresas/{id} {
      allow read: if isSignedIn();
      allow write: if isAdmin();
    }

    match /partidasContables/{id} {
      allow read: if isSignedIn();
      allow write: if isAdmin();
    }

    match /proyectos/{id} {
      allow read: if isSignedIn();
      allow write: if isAdmin();
    }

    match /areas/{id} {
      allow read: if isSignedIn();
      allow write: if isAdmin();
    }

    // ═══════════════════════════════════
    //  PLANILLA WEB
    // ═══════════════════════════════════

    match /history/{id} {
      allow read: if isSignedIn() && (
        resource.data.userId == request.auth.uid
        || isAdmin()
        || isApproverForArea(resource.data.areaId)
      );
      allow create: if isSignedIn() && (
        isAdmin() || request.resource.data.userId == request.auth.uid
      );
      allow update: if isSignedIn() && (
        (isAdmin() && comprobantesWriteOk())
        || (isApproverForArea(resource.data.areaId) && onlyVbFieldsChanged())
      );
      allow delete: if isAdmin();
    }

    match /counters/{dni} {
      allow read, write: if isSignedIn();
    }

    // ═══════════════════════════════════
    //  CAJA CHICA WEB
    // ═══════════════════════════════════

    // Cajas activas (1 por usuario)
    match /cajasChicasActivas/{userId} {
      allow read: if isSignedIn() && (isSelf(userId) || isAdmin());
      allow write: if isSignedIn() && (isSelf(userId) || isAdmin());

      match /expenses/{expenseId} {
        allow read: if isSignedIn() && (isSelf(userId) || isAdmin());
        allow write: if isSignedIn() && (isSelf(userId) || isAdmin());
      }
    }

    // Historial de cajas cerradas
    match /cajasChicasHistorial/{docId} {
      allow read: if isSignedIn();
      allow write: if isSignedIn();

      match /expenses/{expenseId} {
        allow read: if isSignedIn();
        allow write: if isSignedIn();
      }
    }

    // Contadores por proyecto (numeración independiente por obra)
    match /cajasChicasContadores/{obraId} {
      allow read, write: if isSignedIn();
    }
  }
}
```

---

## Storage Rules

1. En Firebase Console → **Storage** → pestaña **Reglas**
2. Reemplaza **todo el contenido** con el bloque de abajo
3. Clic en **Publicar**

```
rules_version = '2';
service firebase.storage {
  match /b/{bucket}/o {

    function esAdmin() {
      return request.auth != null &&
        firestore.get(
          /databases/(default)/documents/users/$(request.auth.uid)
        ).data.rol == 'ADMIN';
    }

    // Consentimientos (Planilla)
    match /consentimientos/{uid}/{filename} {
      allow read: if request.auth != null && (request.auth.uid == uid || esAdmin());
      allow write: if request.auth != null && request.auth.uid == uid;
    }

    // Planillas PDF
    match /planillas/{uid}/{filename} {
      allow read: if request.auth != null && (request.auth.uid == uid || esAdmin());
      allow write: if request.auth != null && (request.auth.uid == uid || esAdmin());
    }

    // Comprobantes de pago (Planilla)
    match /comprobantes/{uid}/{planillaKey}/{filename} {
      allow read: if request.auth != null && (request.auth.uid == uid || esAdmin());
      allow write: if request.auth != null && esAdmin();
    }

    // Comprobantes de Caja Chica (PDF, Imágenes)
    match /caja_chica_comprobantes/{userId}/{allPaths=**} {
      allow read: if request.auth != null;
      allow write: if request.auth != null && (request.auth.uid == userId || esAdmin());
    }

    // Logos de empresa
    match /logos/{allPaths=**} {
      allow read: if request.auth != null;
      allow write: if request.auth != null && esAdmin();
    }
  }
}
```

---

> ✅ **Importante:** Estas reglas mantienen TODAS las reglas originales de Planilla intactas  
> y solo AGREGAN las de Caja Chica al final de cada sección.
