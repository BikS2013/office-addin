import React, { useCallback } from "react";
import { Textarea, makeStyles, tokens } from "@fluentui/react-components";
import { useAppState } from "../context/AppContext";

const useStyles = makeStyles({
  textarea: {
    width: "100%",
    minHeight: "60px",
    marginTop: tokens.spacingVerticalXS,
  },
});

export const PromptInput: React.FC = () => {
  const styles = useStyles();
  const { state, dispatch } = useAppState();

  const handleChange = useCallback(
    (_e: React.ChangeEvent<HTMLTextAreaElement>, data: { value: string }) => {
      dispatch({ type: "SET_PROMPT", prompt: data.value });
    },
    [dispatch]
  );

  return (
    <Textarea
      className={styles.textarea}
      placeholder="Enter prompt (optional)..."
      value={state.prompt}
      onChange={handleChange}
      resize="vertical"
      size="small"
    />
  );
};
