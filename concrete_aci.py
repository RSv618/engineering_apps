import math
from utils import interpolate_linear

class ACIMixDesign:
    def __init__(self):
        # --- CONSTANTS ---
        self.WATER_UNIT_WEIGHT = 62.4  # lb/ft3

        # --- INPUTS (To be defined by user) ---
        self.fc = 0  # Specified compressive strength (psi)
        self.cement_sg = 3.15  # Standard Specific Gravity for Portland Cement
        self.standard_deviation = None  # Psi (None if no data available)
        self.slump_target = 0.0  # Inches (Float allowed)
        self.nmas = 0.0  # Nominal Max Aggregate Size (inches)
        self.is_air_entrained = False  # True/False
        self.exposure_classes = []  # Example: ['F2', 'S1'] or ['W2'] or [] for none

        # Aggregate Properties
        self.ca_sg_ssd = 0.0  # Coarse Agg Specific Gravity (SSD)
        self.ca_absorption = 0.0  # Coarse Agg Absorption (%)
        self.ca_druw = 0.0  # Coarse Agg Dry-Rodded Unit Weight (lb/ft3)
        self.ca_moisture = 0.0  # Coarse Agg Moisture Content (%)
        self.ca_shape = "Angular"  # "Angular" or "Rounded"

        self.fa_sg_ssd = 0.0  # Fine Agg Specific Gravity (SSD)
        self.fa_absorption = 0.0  # Fine Agg Absorption (%)
        self.fa_fineness_modulus = 0.0  # Fineness Modulus
        self.fa_moisture = 0.0  # Fine Agg Moisture Content (%)

        # --- DATA TABLES ---
        self._init_tables()

    def _init_tables(self):
        # Table 5.3.3: Approximate Mixing Water (lb/yd3)
        # REVISED STRATEGY: Keys are the MIDPOINTS of the slump ranges.
        # 1-2" -> 1.5
        # 3-4" -> 3.5
        # 5-6" -> 5.5
        # 6-7" -> 6.5
        self.TABLE_5_3_3_WATER = {
            'Non-Air-Entrained': [
                # (Slump Midpoint, {NMAS: Water})
                (1.5, {0.375: 350, 0.5: 335, 0.75: 315, 1.0: 300, 1.5: 275, 2.0: 260, 3.0: 220}),
                (3.5, {0.375: 385, 0.5: 365, 0.75: 340, 1.0: 325, 1.5: 300, 2.0: 285, 3.0: 245}),
                (5.5, {0.375: 400, 0.5: 375, 0.75: 350, 1.0: 330, 1.5: 305, 2.0: 290, 3.0: 255}),
                (6.5, {0.375: 410, 0.5: 385, 0.75: 360, 1.0: 340, 1.5: 315, 2.0: 300, 3.0: 270}),
            ],
            'Air-Entrained': [
                (1.5, {0.375: 305, 0.5: 295, 0.75: 280, 1.0: 270, 1.5: 250, 2.0: 240, 3.0: 205}),
                (3.5, {0.375: 340, 0.5: 325, 0.75: 305, 1.0: 295, 1.5: 275, 2.0: 265, 3.0: 225}),
                (5.5, {0.375: 355, 0.5: 335, 0.75: 315, 1.0: 300, 1.5: 280, 2.0: 270, 3.0: 240}),
                (6.5, {0.375: 365, 0.5: 345, 0.75: 325, 1.0: 310, 1.5: 290, 2.0: 280, 3.0: 260}),
            ]
        }

        # Table 5.3.3 (Bottom Part): Approximate Air Content (%)
        self.TABLE_5_3_3_AIR = {
            'Entrapped': {0.375: 3.0, 0.5: 2.5, 0.75: 2.0, 1.0: 1.5, 1.5: 1.0, 2.0: 0.5, 3.0: 0.3},
            'F1': {0.375: 6.0, 0.5: 5.5, 0.75: 5.0, 1.0: 4.5, 1.5: 4.5, 2.0: 4.0, 3.0: 3.5},  # Moderate
            'F2_F3': {0.375: 7.5, 0.5: 7.0, 0.75: 6.0, 1.0: 6.0, 1.5: 5.5, 2.0: 5.0, 3.0: 4.5},  # Severe
        }

        # Table 5.3.4: w/cm vs Strength (psi)
        self.TABLE_5_3_4 = [
            (7000, 0.34, 0.32),
            (6000, 0.41, 0.33),
            (5000, 0.48, 0.40),
            (4000, 0.57, 0.48),
            (3000, 0.68, 0.59),
            (2000, 0.82, 0.74)
        ]

        # Table 5.3.6: Bulk Volume of Coarse Aggregate per Unit Volume of Concrete
        self.TABLE_5_3_6 = {
            0.375: {2.4: 0.50, 2.6: 0.48, 2.8: 0.46, 3.0: 0.44},
            0.5: {2.4: 0.59, 2.6: 0.57, 2.8: 0.55, 3.0: 0.53},
            0.75: {2.4: 0.66, 2.6: 0.64, 2.8: 0.62, 3.0: 0.60},
            1.0: {2.4: 0.71, 2.6: 0.69, 2.8: 0.67, 3.0: 0.65},
            1.5: {2.4: 0.75, 2.6: 0.73, 2.8: 0.71, 3.0: 0.69},
            2.0: {2.4: 0.78, 2.6: 0.76, 2.8: 0.74, 3.0: 0.72},
            3.0: {2.4: 0.82, 2.6: 0.80, 2.8: 0.78, 3.0: 0.76}
        }

        # CONSOLIDATED DATA FROM TABLES 4.7.3a, b, c, d (Pages 10-11)
        # Max w/cm limits ('NA' is set to 1.0 effectively meaning no limit)
        self.DURABILITY_LIMITS = {
            'F0': 1.0, 'F1': 0.55, 'F2': 0.45, 'F3': 0.40,
            'S0': 1.0, 'S1': 0.50, 'S2': 0.45, 'S3': 0.45,  # Note: S3 is 0.45 usually, unless special cement
            'W0': 1.0, 'W1': 1.0, 'W2': 0.50,
            'C0': 1.0, 'C1': 1.0, 'C2': 0.40
        }

        # Min specified strength limits ('NA' is set to 2500 or 0)
        # Used to warn user if their design strength is too low for the durability class
        self.DURABILITY_STRENGTH = {
            'F0': 2500, 'F1': 3500, 'F2': 4500, 'F3': 5000,
            'S0': 2500, 'S1': 4000, 'S2': 4500, 'S3': 4500,
            'W0': 2500, 'W1': 2500, 'W2': 4000,
            'C0': 2500, 'C1': 2500, 'C2': 5000
        }

    def calculate_mix(self):
        print("\n--- STARTING ACI 211.1-22 MIX DESIGN ---")

        # --- Step 1: Required Strength ---
        f_cr = self._calculate_f_cr()
        print(f"1. Specified f'c: {self.fc} psi")
        print(f"   Required f'cr: {f_cr} psi")

        # --- Step 2: Verify NMAS ---
        # Just verifying it exists in our lookup
        test_dict = self.TABLE_5_3_3_WATER['Non-Air-Entrained'][0][1]
        if self.nmas not in test_dict:
            raise ValueError(f"NMAS {self.nmas} not supported (must be 0.375, 0.5, 0.75, 1.0, 1.5, 2.0, 3.0)")
        print(f"2. Nominal Max Aggregate Size: {self.nmas} inches")

        # --- Step 3: Water and Air ---
        water_weight, air_percent = self._estimate_water_and_air()

        # Adjustment for Rounded Aggregate (Table 5.3.3.1)
        if self.ca_shape == "Rounded":
            # Reducing water based on Table 5.3.3.1 percentage adjustment
            # "Rounded aggregate: -8%" (approx 25-30 lbs)
            reduction = water_weight * 0.08
            water_weight -= reduction
            print(f"   * Rounded Aggregate Adjustment: -{reduction:.1f} lb")

        print(f"3. Target Slump: {self.slump_target} inches")
        print(f"   Estimated Mixing Water: {water_weight:.1f} lb/yd3")
        print(f"   Target Air Content: {air_percent}%")

        # --- Step 4: w/cm Ratio ---
        wcm = self._select_wcm(f_cr)
        print(f"4. Selected w/cm Ratio: {wcm:.3f}")

        # --- Step 5: Cement Content ---
        cement_weight = water_weight / wcm
        print(f"5. Calculated Cement Content: {cement_weight:.1f} lb/yd3")

        # --- Step 6: Coarse Agg Content ---
        ca_dry_weight = self._estimate_coarse_aggregate()
        print(f"6. Estimated Coarse Aggregate (Dry): {ca_dry_weight:.1f} lb/yd3")

        # --- Step 7: Fine Agg Content (Absolute Volume) ---
        # 7a. Volume calculations
        vol_water = water_weight / self.WATER_UNIT_WEIGHT
        vol_cement = cement_weight / (self.cement_sg * self.WATER_UNIT_WEIGHT)
        vol_air = 27.0 * (air_percent / 100.0)

        # Convert CA Dry Weight to CA SSD Weight for volume calc
        # Vol = Weight_SSD / SG_SSD
        # Weight_SSD = Weight_Dry * (1 + Absorption)
        ca_ssd_weight = ca_dry_weight * (1 + (self.ca_absorption / 100.0))
        vol_ca_ssd = ca_ssd_weight / (self.ca_sg_ssd * self.WATER_UNIT_WEIGHT)

        # 7b. Solve for Sand Volume
        total_vol_minus_sand = vol_water + vol_cement + vol_air + vol_ca_ssd
        vol_sand = 27.0 - total_vol_minus_sand

        # 7c. Convert Sand Volume to Weight (SSD)
        fa_ssd_weight = vol_sand * self.fa_sg_ssd * self.WATER_UNIT_WEIGHT

        print(f"7. Absolute Volumes (ft3):")
        print(f"   Water: {vol_water:.2f}, Cement: {vol_cement:.2f}, Air: {vol_air:.2f}, CA: {vol_ca_ssd:.2f}")
        print(f"   Required Sand Volume: {vol_sand:.2f} ft3")
        print(f"   Fine Aggregate (SSD): {fa_ssd_weight:.1f} lb/yd3")

        # --- Step 8: Moisture Adjustments (Field Weights) ---
        # Adjusting for water ON the aggregates vs ABSORBED by aggregates

        # 1. Determine Oven Dry weights (Base for moisture application)
        ca_od = ca_ssd_weight / (1 + self.ca_absorption / 100.0)
        fa_od = fa_ssd_weight / (1 + self.fa_absorption / 100.0)

        # 2. Calculate Total Water present in the wet aggregates
        # (Weight OD * Total Moisture %)
        total_water_in_ca = ca_od * (self.ca_moisture / 100.0)
        total_water_in_fa = fa_od * (self.fa_moisture / 100.0)

        # 3. Calculate Batch Weights of Wet Aggregates (Solid + Water)
        ca_batch_weight = ca_od + total_water_in_ca
        fa_batch_weight = fa_od + total_water_in_fa

        # 4. Calculate Free Water (Water available to mix)
        absorbed_water_ca = ca_od * (self.ca_absorption / 100.0)
        absorbed_water_fa = fa_od * (self.fa_absorption / 100.0)
        free_water_ca = total_water_in_ca - absorbed_water_ca
        free_water_fa = total_water_in_fa - absorbed_water_fa

        final_batch_water = water_weight - free_water_ca - free_water_fa

        return {
            'f_cr': f_cr,
            'wcm': wcm,
            'air_percent': air_percent,
            # Imperial Weights (lbs) per 1 cubic yard
            'weights_lb': {
                'cement': cement_weight,
                'water_net': final_batch_water,
                'ca_wet': ca_batch_weight,
                'fa_wet': fa_batch_weight,
                'total': cement_weight + final_batch_water + ca_batch_weight + fa_batch_weight
            },
            # Absolute Volumes (ft3) per 1 cubic yard
            'volumes_ft3': {
                'cement': vol_cement,
                'water': vol_water,
                'ca': vol_ca_ssd,
                'fa': vol_sand,
                'air': vol_air
            }
        }

    def _calculate_f_cr(self):
        if self.standard_deviation is None:
            if self.fc < 3000:
                return self.fc + 1000
            elif self.fc <= 5000:
                return self.fc + 1200
            else:
                return 1.1 * self.fc + 700
        else:
            s = self.standard_deviation
            if self.fc <= 5000:
                return max(self.fc + 1.34 * s, self.fc + 2.33 * s - 500)
            else:
                return max(self.fc + 1.34 * s, 0.90 * self.fc + 2.33 * s)

    def _estimate_water_and_air(self):
        type_key = 'Air-Entrained' if self.is_air_entrained else 'Non-Air-Entrained'
        data_points = self.TABLE_5_3_3_WATER[type_key]

        slump = self.slump_target
        nmas = self.nmas

        # -- INTERPOLATION LOGIC FOR WATER --
        # We have points at 1.5, 3.5, 5.5, 6.5.
        # We need to find the two points bounding our target slump.

        x1, y1, x2, y2 = None, None, None, None

        # Handle out of bounds (Clamp or Error? Let's clamp for safety but warn)
        if slump < 1.5:
            # Extrapolate or clamp to min? Clamping to 1.5 value is safer for code stability
            x1, y1 = 1.5, data_points[0][1][nmas]
            water = y1
        elif slump > 6.5:
            # Clamp to max
            water = data_points[-1][1][nmas]
        else:
            # Find the bracket
            for i in range(len(data_points) - 1):
                curr_slump = data_points[i][0]
                next_slump = data_points[i + 1][0]

                if curr_slump <= slump <= next_slump:
                    x1 = curr_slump
                    y1 = data_points[i][1][nmas]
                    x2 = next_slump
                    y2 = data_points[i + 1][1][nmas]
                    break

            # Interpolate
            water = interpolate_linear(slump, x1, y1, x2, y2)

        # -- AIR CONTENT --
        if self.is_air_entrained:
            # Check if any severe class (F2 or F3) is in the active exposures
            if 'F2' in self.exposure_classes or 'F3' in self.exposure_classes:
                air = self.TABLE_5_3_3_AIR['F2_F3'][self.nmas]
            elif 'F1' in self.exposure_classes:
                air = self.TABLE_5_3_3_AIR['F1'][self.nmas]
            else:
                # If air entrained is requested but no specific F-class is listed,
                # standard practice is to default to Moderate (F1) table values
                # or just use the table provided in 5.3.3 for "Air-Entrained"
                air = self.TABLE_5_3_3_AIR['F1'][self.nmas]
        else:
            air = self.TABLE_5_3_3_AIR['Entrapped'][self.nmas]

        return water, air

    def _select_wcm(self, f_cr):
        # 1. Strength-based w/cm (Interpolation)
        table = sorted(self.TABLE_5_3_4, key=lambda x: x[0], reverse=True)  # Descending strength
        idx_wcm = 2 if self.is_air_entrained else 1

        strength_wcm = 0.0

        if f_cr > table[0][0]:
            # Strength higher than table max (7000/6000).
            # Standard ACI 211 doesn't cover high-strength concrete (ACI 211.4R does).
            # We clamp to the lowest available w/cm in the table.
            strength_wcm = table[0][idx_wcm]
            print(f"   WARNING: Required f'cr ({f_cr}) exceeds table data. Using min w/cm.")
        elif f_cr < table[-1][0]:
            # Strength lower than table min (2000). Use max w/cm.
            strength_wcm = table[-1][idx_wcm]
        else:
            for i in range(len(table) - 1):
                s_high, w_high = table[i][0], table[i][idx_wcm]
                s_low, w_low = table[i + 1][0], table[i + 1][idx_wcm]

                if s_low <= f_cr <= s_high:
                    # Linear interpolation
                    strength_wcm = interpolate_linear(f_cr, s_low, w_low, s_high, w_high)
                    break

        print(f"   - w/cm for Strength ({f_cr:.0f} psi): {strength_wcm:.3f}")

        # --- 2. CALCULATE DURABILITY-BASED w/cm (The Edit) ---

        # Default to no restriction (1.0 w/cm) and 0 psi requirement
        most_restrictive_wcm = 1.0
        highest_req_fc = 0
        governing_class_wcm = "None"
        governing_class_fc = "None"

        # Iterate over the list of exposure classes (e.g., ['F2', 'S1'])
        if self.exposure_classes:
            for exp in self.exposure_classes:
                # Look up w/cm limit
                limit = self.DURABILITY_LIMITS.get(exp, 1.0)
                if limit < most_restrictive_wcm:
                    most_restrictive_wcm = limit
                    governing_class_wcm = exp

                # Look up Min f'c limit
                min_fc = self.DURABILITY_STRENGTH.get(exp, 0)
                if min_fc > highest_req_fc:
                    highest_req_fc = min_fc
                    governing_class_fc = exp

            print(f"   - Durability Check: Governing Class {governing_class_wcm} "
                  f"(Max w/cm: {most_restrictive_wcm})")
            print(f"   - Durability Check: Governing Class {governing_class_fc} "
                  f"(Min f'c: {highest_req_fc} psi)")

            # Vital Check: Does durability require higher strength than the structural design?
            if self.fc < highest_req_fc:
                print(f"   CRITICAL WARNING: Your specified f'c ({self.fc} psi) is LOWER "
                      f"than the durability requirement for {governing_class_fc} ({highest_req_fc} psi). "
                      f"You must increase your specified strength.")
        else:
            print("   - No exposure classes defined.")

        # --- 3. DETERMINE FINAL GOVERNING w/cm ---
        final_wcm = min(strength_wcm, most_restrictive_wcm)

        if final_wcm == most_restrictive_wcm:
            print(f"   * GOVERNS: Durability ({governing_class_wcm})")
        else:
            print(f"   * GOVERNS: Strength (f'cr = {f_cr:.0f})")

        return final_wcm

    def _estimate_coarse_aggregate(self):
        # Table 5.3.6 Interpolation
        row = self.TABLE_5_3_6.get(self.nmas)
        if not row: raise ValueError("NMAS invalid for CA table")

        fm = self.fa_fineness_modulus

        # Known FM columns: 2.4, 2.6, 2.8, 3.0
        # Interpolate if between columns
        if fm <= 2.4:
            vol_frac = row[2.4]
        elif fm >= 3.0:
            vol_frac = row[3.0]
        else:
            # Find immediate neighbors
            keys = [2.4, 2.6, 2.8, 3.0]
            for i in range(len(keys) - 1):
                if keys[i] <= fm <= keys[i + 1]:
                    v1 = row[keys[i]]
                    v2 = row[keys[i + 1]]
                    vol_frac = interpolate_linear(fm, keys[i], v1, keys[i + 1], v2)
                    break
            else:
                raise ValueError(f'Cannot find key {fm} in TABLE 5.3.6')

        print(f"   - Coarse Agg Volume Fraction (b/b0): {vol_frac:.3f}")
        return vol_frac * self.ca_druw * 27.0


if __name__ == "__main__":
    aci_mix = ACIMixDesign()

    # --- TEST CASE ---
    # ACI 211.1 Example 9.2 inputs
    aci_mix.fc = 2500
    aci_mix.standard_deviation = None
    aci_mix.slump_target = 3.5

    aci_mix.nmas = 1.5
    aci_mix.is_air_entrained = False
    aci_mix.exposure_classes = ['F0', 'S0', 'W0', 'C0']

    aci_mix.ca_sg_ssd = 2.68
    aci_mix.ca_druw = 100.0
    aci_mix.ca_moisture = 2.0
    aci_mix.ca_shape = "Rounded"
    aci_mix.ca_absorption = 0.5

    aci_mix.fa_fineness_modulus = 2.80
    aci_mix.fa_sg_ssd = 2.64
    aci_mix.fa_absorption = 0.7
    aci_mix.fa_moisture = 6.0

    aci_mix.run_design()