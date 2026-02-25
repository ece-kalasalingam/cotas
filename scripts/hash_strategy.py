import hashlib

class HashStrategy:
    """
    Single Source of Truth for template fingerprinting.
    Used to 'lock' the template during generation and 'verify' it during upload.
    """
    @staticmethod
    def compute_structure_hash(validated_setup, canon_func) -> str:
        hasher = hashlib.sha256()

        # 1. Components & Questions (Deterministic Sort)
        for name, comp in sorted(validated_setup.components.items()):
            hasher.update(name.encode())
            for q in comp.questions:
                hasher.update(str(q.identifier).encode())
                # Use fixed-point for floats to avoid precision noise
                hasher.update(f"{float(q.max_marks):.2f}".encode())
                # Apply the strict CO canonicalizer
                hasher.update(canon_func(",".join(q.co_list)).encode())

        # 2. Students (Deterministic Sort)
        student_regnos = [str(s.reg_no).strip().upper() for s in validated_setup.students]
        for regno in sorted(student_regnos):
            hasher.update(regno.encode())

        return hasher.hexdigest()