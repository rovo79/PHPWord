<?php

require 'vendor/autoload.php';

use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\Element\Text;
use PhpOffice\PhpWord\Element\TextRun;
use PhpOffice\PhpWord\Element\ListItem;
use PhpOffice\PhpWord\Element\ListItemRun;
use PhpOffice\PhpWord\Element\Table;
use PhpOffice\PhpWord\Element\Title;

function parseDocx($filePath)
{
    $phpWord = IOFactory::load($filePath);
    $parsedContent = [];

    foreach ($phpWord->getSections() as $section) {
        foreach ($section->getElements() as $element) {
            if ($element instanceof Text) {
                $parsedContent[] = [
                    'type' => 'text',
                    'content' => $element->getText(),
                    'style' => $element->getFontStyle()
                ];
            } elseif ($element instanceof TextRun) {
                $textRunContent = [];
                foreach ($element->getElements() as $textElement) {
                    if ($textElement instanceof Text) {
                        $textRunContent[] = [
                            'type' => 'text',
                            'content' => $textElement->getText(),
                            'style' => $textElement->getFontStyle()
                        ];
                    }
                }
                $parsedContent[] = [
                    'type' => 'textrun',
                    'content' => $textRunContent
                ];
            } elseif ($element instanceof ListItem) {
                $parsedContent[] = [
                    'type' => 'listitem',
                    'content' => $element->getText(),
                    'depth' => $element->getDepth(),
                    'style' => $element->getStyle()
                ];
            } elseif ($element instanceof ListItemRun) {
                $listItemRunContent = [];
                foreach ($element->getElements() as $listItemElement) {
                    if ($listItemElement instanceof Text) {
                        $listItemRunContent[] = [
                            'type' => 'text',
                            'content' => $listItemElement->getText(),
                            'style' => $listItemElement->getFontStyle()
                        ];
                    }
                }
                $parsedContent[] = [
                    'type' => 'listitemrun',
                    'content' => $listItemRunContent,
                    'depth' => $element->getDepth(),
                    'style' => $element->getStyle()
                ];
            } elseif ($element instanceof Table) {
                $tableContent = [];
                foreach ($element->getRows() as $row) {
                    $rowContent = [];
                    foreach ($row->getCells() as $cell) {
                        $cellContent = [];
                        foreach ($cell->getElements() as $cellElement) {
                            if ($cellElement instanceof Text) {
                                $cellContent[] = [
                                    'type' => 'text',
                                    'content' => $cellElement->getText(),
                                    'style' => $cellElement->getFontStyle()
                                ];
                            }
                        }
                        $rowContent[] = $cellContent;
                    }
                    $tableContent[] = $rowContent;
                }
                $parsedContent[] = [
                    'type' => 'table',
                    'content' => $tableContent,
                    'style' => $element->getStyle()
                ];
            } elseif ($element instanceof Title) {
                $parsedContent[] = [
                    'type' => 'title',
                    'content' => $element->getText(),
                    'depth' => $element->getDepth(),
                    'style' => $element->getStyle()
                ];
            }
        }
    }

    return $parsedContent;
}

function convertToHtml($parsedContent)
{
    $htmlContent = '';

    foreach ($parsedContent as $element) {
        switch ($element['type']) {
            case 'text':
                $htmlContent .= '<p>' . htmlspecialchars($element['content']) . '</p>';
                break;
            case 'textrun':
                $htmlContent .= '<p>';
                foreach ($element['content'] as $textElement) {
                    $htmlContent .= htmlspecialchars($textElement['content']);
                }
                $htmlContent .= '</p>';
                break;
            case 'listitem':
                $htmlContent .= str_repeat('<ul>', $element['depth']);
                $htmlContent .= '<li>' . htmlspecialchars($element['content']) . '</li>';
                $htmlContent .= str_repeat('</ul>', $element['depth']);
                break;
            case 'listitemrun':
                $htmlContent .= str_repeat('<ul>', $element['depth']);
                $htmlContent .= '<li>';
                foreach ($element['content'] as $listItemElement) {
                    $htmlContent .= htmlspecialchars($listItemElement['content']);
                }
                $htmlContent .= '</li>';
                $htmlContent .= str_repeat('</ul>', $element['depth']);
                break;
            case 'table':
                $htmlContent .= '<table>';
                foreach ($element['content'] as $row) {
                    $htmlContent .= '<tr>';
                    foreach ($row as $cell) {
                        $htmlContent .= '<td>';
                        foreach ($cell as $cellElement) {
                            $htmlContent .= htmlspecialchars($cellElement['content']);
                        }
                        $htmlContent .= '</td>';
                    }
                    $htmlContent .= '</tr>';
                }
                $htmlContent .= '</table>';
                break;
            case 'title':
                $htmlContent .= '<h' . $element['depth'] . '>' . htmlspecialchars($element['content']) . '</h' . $element['depth'] . '>';
                break;
        }
    }

    return $htmlContent;
}

if ($argc < 2) {
    echo "Usage: php cli.php <path_to_docx_file>\n";
    exit(1);
}

$filePath = $argv[1];
$parsedContent = parseDocx($filePath);
$htmlContent = convertToHtml($parsedContent);

echo $htmlContent;
