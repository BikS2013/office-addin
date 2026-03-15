import React from "react";
import {
  Accordion,
  AccordionItem,
  AccordionHeader,
  AccordionPanel,
  Text,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import type { ApiGroup as ApiGroupType } from "../types";
import { ApiItem } from "./ApiItem";

interface ApiGroupProps {
  readonly group: ApiGroupType;
}

const useStyles = makeStyles({
  description: {
    color: tokens.colorNeutralForeground3,
    display: "block",
    padding: `0 ${tokens.spacingHorizontalM}`,
    marginBottom: tokens.spacingVerticalXS,
  },
  panel: {
    paddingLeft: tokens.spacingHorizontalM,
  },
});

export const ApiGroup: React.FC<ApiGroupProps> = ({ group }) => {
  const styles = useStyles();

  const hasNestedGroups = group.groups && group.groups.length > 0;
  const hasApis = group.apis && group.apis.length > 0;

  return (
    <AccordionItem value={group.id}>
      <AccordionHeader size="small">
        <Text weight="semibold" size={300}>
          {group.name}
        </Text>
      </AccordionHeader>
      <AccordionPanel className={styles.panel}>
        {group.description && (
          <Text className={styles.description} size={200}>
            {group.description}
          </Text>
        )}

        {hasNestedGroups && (
          <Accordion collapsible multiple>
            {group.groups!.map((nestedGroup) => (
              <ApiGroup key={nestedGroup.id} group={nestedGroup} />
            ))}
          </Accordion>
        )}

        {hasApis &&
          group.apis!.map((api) => (
            <ApiItem key={api.id} api={api} />
          ))}
      </AccordionPanel>
    </AccordionItem>
  );
};
