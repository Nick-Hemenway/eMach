from FEA_ana import FEAProblem, RotationCondition
from jmag_solver import JMagTransient2DFEA
import mach_cad.model_obj as mo

# Create cross-sections

# stator cross-section
stator1 = mo.CrossSectInnerRotorStator(
    name="Stator",
    dim_alpha_st=mo.DimDegree(44.5),
    dim_alpha_so=mo.DimDegree((44.5 / 2)),
    dim_r_si=mo.DimMillimeter(14.16),
    dim_d_sy=mo.DimMillimeter(13.54),
    dim_d_st=mo.DimMillimeter(16.94),
    dim_d_sp=mo.DimMillimeter(8.14),
    dim_d_so=mo.DimMillimeter(5.43),
    dim_w_st=mo.DimMillimeter(9.1),
    dim_r_st=mo.DimMillimeter(0),
    dim_r_sf=mo.DimMillimeter(0),
    dim_r_sb=mo.DimMillimeter(0),
    Q=6,
    location=mo.Location2D(anchor_xy=[mo.DimMillimeter(0), mo.DimMillimeter(0)]),
    theta=mo.DimDegree(0),
)

# rotor cross-section
rotor1 = mo.CrossSectInnerNotchedRotor(
    name="Shaft",
    location=mo.Location2D(),
    dim_alpha_rm=mo.DimDegree(180),
    dim_alpha_rs=mo.DimDegree(90),
    dim_d_ri=mo.DimMillimeter(8),
    dim_r_ri=mo.DimMillimeter(0),
    dim_d_rp=mo.DimMillimeter(5),
    dim_d_rs=mo.DimMillimeter(3),
    p=1,
    s=2,
)

# all magnet cross-sections
magnets = []
for i in range(2):
    magnet = mo.CrossSectArc(
        name="Magnet" + str(i),
        location=mo.Location2D(theta=mo.DimDegree(180 * i)),
        dim_d_a=mo.DimMillimeter(3.41),
        dim_r_o=mo.DimMillimeter(11.41),
        dim_alpha=mo.DimDegree(180),
    )
    magnets.append(magnet)

# example coil cross-section
coils = []
Q = 6
for slot in range(Q):
    for layer in range(2):
        coil = mo.CrossSectInnerRotorStatorCoil(
            name="Slot%d_%d" % (slot + 1, layer + 1),
            dim_r_si=mo.DimMillimeter(14.16),
            dim_d_st=mo.DimMillimeter(16.94),
            dim_d_sp=mo.DimMillimeter(8.14),
            dim_w_st=mo.DimMillimeter(9.1),
            dim_r_st=mo.DimMillimeter(0),
            dim_r_sf=mo.DimMillimeter(0),
            dim_r_sb=mo.DimMillimeter(0),
            Q=6,
            slot=slot + 1,
            layer=layer,
            num_of_layers=2,
            location=mo.Location2D(
                anchor_xy=[mo.DimMillimeter(0), mo.DimMillimeter(0)]
            ),
            theta=mo.DimDegree(0),
        )
        coils.append(coil)

# Create components

StatorComp = mo.Component(
    name=stator1.name,
    cross_sections=[stator1],
    material=mo.MaterialGeneric(name="10JNEX900", color=r"#808080"),
    make_solid=mo.MakeExtrude(location=mo.Location3D(), dim_depth=mo.DimMillimeter(25)),
)

RotorComp = mo.Component(
    name=rotor1.name,
    cross_sections=[rotor1],
    material=mo.MaterialGeneric(name="10JNEX900", color=r"#808080"),
    make_solid=mo.MakeExtrude(location=mo.Location3D(), dim_depth=mo.DimMillimeter(25)),
)

MagnetComps = []
for i in range(len(magnets)):
    magnet_comp = mo.Component(
        name=magnets[i].name,
        cross_sections=[magnets[i]],
        material=mo.MaterialGeneric(name="Arnold/Reversible/N40H", color=r"#4d4b4f"),
        make_solid=mo.MakeExtrude(
            location=mo.Location3D(), dim_depth=mo.DimMillimeter(25)
        ),
    )
    MagnetComps.append(magnet_comp)

CoilComps = []
for i in range(len(coils)):
    coil_comp = mo.Component(
        name=coils[i].name,
        cross_sections=[coils[i]],
        material=mo.MaterialGeneric(name="Copper", color=r"#b87333"),
        make_solid=mo.MakeExtrude(
            location=mo.Location3D(), dim_depth=mo.DimMillimeter(25)
        ),
    )
    CoilComps.append(coil_comp)

# Setup conditions

rot_cond = RotationCondition(["Shaft", "Magnet0", "Magnet1"], cond_para=[1000, 0])

# Create problem

components = (
    StatorComp,
    RotorComp,
    MagnetComps[0],
    MagnetComps[1],
    CoilComps[0],
    CoilComps[1],
    CoilComps[2],
    CoilComps[3],
    CoilComps[4],
    CoilComps[5],
    CoilComps[6],
    CoilComps[7],
    CoilComps[8],
    CoilComps[9],
    CoilComps[10],
    CoilComps[11],
)

prob = FEAProblem(
    components=components,
    conditions=[rot_cond],
    settings=None,
    get_results=None,
    config=None,
)
tool = JMagTransient2DFEA()
tool.run(prob)
