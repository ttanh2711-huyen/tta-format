import { checkPartyHeaderFooter } from "../standards/hd05/checks";
import { checkPartyStructure } from "../standards/hd05/structureChecks";
import { checkNd30HeaderFooter } from "../standards/nd30/checks";
import { checkNd30Basis } from "../standards/nd30/decisionChecks";
import { checkNd30Structure } from "../standards/nd30/structureChecks";
import type { StandardProfile } from "../standards/types";
import type { Rule } from "./engine";
import { checkCapitalization } from "./capitalization";
import { checkMixedStandards } from "./mixedStandards";
import { checkPointLabels } from "./pointLabels";
import { checkUrgencyMarks } from "./urgencyRules";

/** Các bộ quy tắc khối đầu và khối cuối theo kiểu khối của chuẩn (`headerKind`), kèm phát hiện tiêu đề của chuẩn khác. */
export function headerFooterRules(profile: StandardProfile): Rule[] {
  return profile.headerKind === "quocHieu" ? [checkNd30HeaderFooter, checkNd30Basis, checkNd30Structure, checkUrgencyMarks, checkMixedStandards, checkCapitalization, checkPointLabels] : [checkPartyHeaderFooter, checkPartyStructure, checkUrgencyMarks, checkMixedStandards, checkCapitalization, checkPointLabels];
}
