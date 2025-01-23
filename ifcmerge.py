import sys
import ifcopenshell
from glob import glob

new_file = ifcopenshell.file(schema='IFC4')


if __name__ == '__main__':
    old_files = list(map(ifcopenshell.open, list(glob(f'{sys.argv[1]}/*.ifc'))))

    for of in old_files:
        for inst in of:
            new_file.add(inst)

    projects = new_file.by_type('IfcProject')

    for p in projects[1:]:
        refs = new_file.get_inverse(p)
        is_project = lambda inst: inst == p
        assign_first = lambda inst: projects[0]

        for ref in refs:
            for idx, attr in enumerate(ref):
                ref[idx] = ifcopenshell.entity_instance.walk(
                    is_project,
                    assign_first,
                    ref[idx]
                )

        new_file.remove(p)

new_file.write(f'{sys.argv[1]}/merged.ifc')