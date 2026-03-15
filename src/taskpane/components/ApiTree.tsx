import React from "react";
import {
  Accordion,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import type { ApiGroup as ApiGroupType } from "../types";
import { ApiGroup } from "./ApiGroup";

interface ApiTreeProps {
  readonly groups: readonly ApiGroupType[];
}

const useStyles = makeStyles({
  container: {
    padding: tokens.spacingVerticalXS,
  },
});

export const ApiTree: React.FC<ApiTreeProps> = ({ groups }) => {
  const styles = useStyles();

  if (groups.length === 0) {
    return null;
  }

  return (
    <div className={styles.container}>
      <Accordion collapsible multiple>
        {groups.map((group) => (
          <ApiGroup key={group.id} group={group} />
        ))}
      </Accordion>
    </div>
  );
};
