# Role Authorization Guide: COPYEDITPM & ADMIN Access

## ✅ Role-Based Access Control

Both **COPYEDITPM** and **ADMIN** roles have **EQUAL FULL ACCESS** to all discovery and rule selection features.

---

## 🔐 Authorized Roles

### **ALLOWED_ROLES Configuration**
```python
# In manuscript_bp.py (line 65):
ALLOWED_ROLES = {'COPYEDIT', 'COPYEDITPM', 'PM', 'ADMIN'}
```

### **Roles with Full Access:**
- ✅ **COPYEDITPM** - Copy editors and project managers
- ✅ **ADMIN** - System administrators
- ✅ **PM** - Project managers (legacy)
- ✅ **COPYEDIT** - Copy editors (legacy)

### **Roles WITHOUT Access:**
- ❌ Any other role (returns 403 Forbidden)
- ❌ Unauthenticated users (redirected to login)

---

## 📍 Protected Routes: Both Roles Allowed

### **Discovery Routes**
```
@manuscript_auth_required
GET  /manuscript/analyze
GET  /manuscript/discovery?session_id=<id>
GET  /manuscript/discovery/<session_id>/ia-rows
POST /manuscript/discovery/<session_id>/create-selection
GET  /manuscript/discovery/<session_id>/ia-report
```

### **Rule Selection Routes**
```
@manuscript_auth_required
GET  /manuscript/rule-selections
POST /manuscript/rule-selections/<id>/activate
DELETE /manuscript/rule-selections/<id>
```

### **Editor Review Routes**
```
@manuscript_auth_required
GET  /manuscript/review/<job_id>
POST /manuscript/review/<job_id>/apply-fixes
```

---

## 🛡️ How Authorization Works

### **Step 1: Login**
```python
# After login, session contains:
session['user_id'] = 'john@company.com'
session['role'] = 'COPYEDITPM'  # or 'ADMIN'
session['is_admin'] = False     # or True (set separately)
```

### **Step 2: Route Protection**
```python
# @manuscript_auth_required decorator (line 68):
def manuscript_auth_required(f):
    """Auth guard for manuscript routes."""
    @wraps(f)
    def wrapped(*args, **kwargs):
        # Check 1: Is user logged in?
        if 'user_id' not in session:
            flash("Please log in to continue.")
            return redirect(url_for('login'))
        
        # Check 2: Does role match allowed list?
        role = (session.get('role') or '').upper()
        if not session.get('is_admin') and role not in ALLOWED_ROLES:
            flash("You don't have permission to access this page.", "error")
            return redirect(url_for('dashboard'))
        
        # ✅ Both checks pass → Allow access
        return f(*args, **kwargs)
    return wrapped
```

### **Step 3: Template Access**
All templates render **identical content** for both COPYEDITPM and ADMIN:
- No role-based UI hiding
- No conditional menus
- Same features available to both

---

## 📋 Feature Access Matrix

| Feature | COPYEDITPM | ADMIN | Notes |
|---------|-----------|-------|-------|
| Upload manuscripts | ✅ | ✅ | Both can upload chapters |
| Select rules | ✅ | ✅ | Both can use discovery UI |
| Save selections | ✅ | ✅ | Both can create selections |
| View selections | ✅ | ✅ | Both can list saved selections |
| Activate selection | ✅ | ✅ | Both can make selection active |
| Edit selection | ✅ | ✅ | Both can modify rules |
| Delete selection | ✅ | ✅ | Both can remove selections |
| Review findings | ✅ | ✅ | Both can use editor review |
| Apply fixes | ✅ | ✅ | Both can apply fixes & download |
| Access analytics | ✅ | ✅ | Both see dashboards |

---

## 🔍 No Differentiation in Functionality

### **What's THE SAME for Both Roles:**

1. **Discovery UI**
   - Same three-panel layout (Elements | Rules | Stats)
   - Same rule selection interface
   - Same live statistics
   - Same save selection form

2. **Rule Selections**
   - Same DataTable display
   - Same Activate/Edit/Delete buttons
   - Same "Create New" modal
   - Same status badges

3. **Editor Review**
   - Same three-panel editor (Finding list | Context | Preview)
   - Same Accept/Reject workflow
   - Same track changes output
   - Same download functionality

### **What's DIFFERENT:**
- **Nothing** - Both roles see and do exactly the same things
- All features are equally accessible
- No admin-only menus or buttons
- No restricted options for COPYEDITPM

---

## 📝 Implementation Details

### **Database Model**
```python
# In manuscript_core/models.py:
class RuleSelection:
    id: int
    session_id: str
    selection_name: str
    selected_ia_rows: JSON
    custom_grouping: JSON
    created_by: str        # ← Stores who created it (not role-restricted)
    created_at: datetime
    active: bool
```

No role field in RuleSelection - access is controlled at route level only.

### **Backend Validation**
```python
# Routes don't check role AGAIN after @manuscript_auth_required
# If you got past the decorator, you're authorized

@manuscript_bp.route('/discovery/<session_id>/create-selection', methods=['POST'])
@manuscript_auth_required
def create_selection(session_id: str):
    # No additional role checks here
    # Both COPYEDITPM and ADMIN can execute this
    selection = RuleSelection(...)
    selection.save(db)
    return jsonify({"selection_id": selection.id})
```

### **Frontend Templates**
```html
<!-- discovery.html, rule_selections.html, editor_review.html -->
<!-- No conditional rendering based on role -->
<!-- All users see identical UI -->
```

---

## 🚀 Usage: Same Workflow for Both Roles

### **COPYEDITPM User Flow:**
```
1. Login as COPYEDITPM
2. Navigate to /manuscript/analyze
3. Upload chapters
4. Go to /manuscript/discovery?session_id=XYZ
5. Select rules, save selection
6. Go to /manuscript/rule-selections
7. Activate selection
8. Go to /manuscript/review/<job_id>
9. Review and apply fixes
```

### **ADMIN User Flow:**
```
1. Login as ADMIN
2. Navigate to /manuscript/analyze
3. Upload chapters
4. Go to /manuscript/discovery?session_id=XYZ
5. Select rules, save selection
6. Go to /manuscript/rule-selections
7. Activate selection
8. Go to /manuscript/review/<job_id>
9. Review and apply fixes
```

**Both flows are IDENTICAL.**

---

## ✨ Why No Differentiation?

**Design Decision:**
- Discovery and rule selection are **core editorial features**
- Both COPYEDITPM and ADMIN need full control over manuscript analysis
- Creating role-based feature splits would complicate the UI
- Admin role doesn't need extra powers in this context
- If admin-only features are needed, they'd be in a separate admin panel

**Result:**
- Simplified codebase
- Consistent user experience
- No confusion about "what can I access?"
- Full autonomy for both roles on manuscript work

---

## 🔐 Security Notes

### **What IS Protected:**
- Access to manuscript features (requires authentication)
- Role validation (must be in ALLOWED_ROLES)
- Data access (only users with valid session can load data)

### **What's NOT Role-Differentiated:**
- Features within manuscript tools (all equally accessible)
- Selection management (both can save/activate/delete)
- Data visibility (both can see all analysis results)

### **Audit Trail:**
```python
# Who did what is tracked via:
RuleSelection.created_by = session.get("username", "unknown")
RuleSelection.created_at = datetime.utcnow()
```

Even if both roles can do everything, the database records who created each selection.

---

## ✅ Verification Checklist

- ✅ COPYEDITPM in ALLOWED_ROLES (line 65, manuscript_bp.py)
- ✅ ADMIN allowed by is_admin flag (line 76, manuscript_bp.py)
- ✅ All discovery routes use @manuscript_auth_required
- ✅ All rule_selections routes use @manuscript_auth_required
- ✅ All editor_review routes use @manuscript_auth_required
- ✅ Templates have no role-conditional rendering
- ✅ Database models don't restrict by role
- ✅ Both roles tested and working (equal access confirmed)

---

## 🎯 Summary

**Both COPYEDITPM and ADMIN roles have:**
- ✅ Full access to all discovery features
- ✅ Full access to all rule selection features
- ✅ Full access to all editor review features
- ✅ Identical user interface
- ✅ Identical functionality
- ✅ Equal permissions

**No feature differentiation between the two roles.**

---

**Last Updated**: 2026-05-09  
**Status**: Both roles fully authorized and tested  
**Authorization Method**: @manuscript_auth_required decorator + ALLOWED_ROLES set
